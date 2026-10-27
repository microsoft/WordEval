"""Layout-coupled interaction generator for interaction-sensitive compositions.

Each case composes geometric operations where operation A changes the spatial
context that operation B depends on. All variants must pass the dependency
test: swapping the order of the two operations produces a different final
state. Variants that are order-independent belong in
``additive_composition_x_object_effect`` instead.

Coupling notes from empirical verification:
- ``rotate → align`` IS coupled because ``ShapeRange.Align`` uses the
  visual (rotated) bounding box.
- ``rotate → distribute`` is NOT coupled because ``ShapeRange.Distribute``
  uses un-rotated ``.Left/.Top`` (moved to additive).
- ``group → flip/rotate → ungroup`` IS coupled because the group's bounding
  box provides spatial context that individual shapes lack.
- ``wrap → position`` IS coupled because inline shapes ignore explicit
  position; the wrap step provides the layout context that makes the
  position step meaningful.
- ``position preset → align-multi`` IS coupled because the align step's
  reference point (rightmost / topmost) changes after one shape is moved
  through a standard Position gallery preset.

Includes 5 groups / 13 variants:
- LayoutCoupledRotateAlign (4 variants) — rotate_then_align only
- LayoutCoupledGroupLifecycle (3 variants)
- LayoutCoupledWrapPosition (2 variants)
- LayoutCoupledMultiStep (2 variants)
- LayoutCoupledPositionAlignMulti (2 variants) — new, framing-aligned

Usage:
    python -m generator.compositions.run --test-set interaction_composition_x_layout_coupled
"""

import json
import time
from datetime import datetime
from pathlib import Path

from generator.config import BENCHMARK_OUTPUT_DIR
from generator.execution.setup import ensure_word
from generator.compositions.atomic_intents import (
    build_layout_coupled_atomic_intents,
)

from .shared import (
    build_spec,
    close_doc,
    create_spread_tall_shapes_doc,
    create_tall_shapes_doc,
    get_shared_input,
    open_doc,
    save_as_gold,
    write_instructions_json,
    write_spec,
)

TEST_SET = 'interaction_composition_x_layout_coupled'


# ═════════════════════════════════════════════════════════════════════════
# Case definitions — grouped by command/directory name
# ═════════════════════════════════════════════════════════════════════════

CASE_GROUPS = [
    # ── Rotation + Alignment (4 variants) ─────────────────────────────
    #
    # Genuinely layout-coupled: ShapeRange.Align uses the visual (rotated)
    # bounding box, so rotating the middle shape first changes where its
    # visual top/bottom/left/right are, which in turn changes the
    # alignment target chosen by the subsequent align step.
    {
        'group': 'LayoutCoupledRotateAlign',
        'target_object_type': 'multi_shape',
        'behavior': 'object_effect',
        'verification_method': 'composition_dispatch',
        'doc_in_name': 'doc_in_three_shapes_tall',
        'needs_tall_shapes': True,
        'variants': {
            'variant_right_top': {
                'steps': [
                    {'operation': 'rotate_shape',
                     'params': {'shape_index': 2, 'angle': 90}},
                    {'operation': 'align_shapes',
                     'params': {'align_type': 3}},  # msoAlignTops
                ],
                'interaction_class': 'rotate_then_align',
                'verification_description':
                    'Middle rectangle rotated 90° clockwise, then all three '
                    'rectangles top-aligned',
                'instruction':
                    "Rotate the middle rectangle 90° clockwise, then "
                    "align all three rectangles by their top edges.",
            },
            'variant_left_left': {
                'steps': [
                    {'operation': 'rotate_shape',
                     'params': {'shape_index': 2, 'angle': -90}},
                    {'operation': 'align_shapes',
                     'params': {'align_type': 0}},  # msoAlignLefts
                ],
                'interaction_class': 'rotate_then_align',
                'verification_description':
                    'Middle rectangle rotated 90° counter-clockwise, then '
                    'all three rectangles left-aligned',
                'instruction':
                    "Rotate the middle rectangle 90° counter-clockwise, then "
                    "align all three rectangles by their left edges.",
            },
            'variant_right_bottom': {
                'steps': [
                    {'operation': 'rotate_shape',
                     'params': {'shape_index': 2, 'angle': 90}},
                    {'operation': 'align_shapes',
                     'params': {'align_type': 5}},  # msoAlignBottoms
                ],
                'interaction_class': 'rotate_then_align',
                'verification_description':
                    'Middle rectangle rotated 90° clockwise, then all three '
                    'rectangles bottom-aligned',
                'instruction':
                    "Rotate the middle rectangle 90° clockwise, then "
                    "align all three rectangles by their bottom edges.",
            },
            'variant_left_right': {
                'steps': [
                    {'operation': 'rotate_shape',
                     'params': {'shape_index': 2, 'angle': -90}},
                    {'operation': 'align_shapes',
                     'params': {'align_type': 2}},  # msoAlignRights
                ],
                'interaction_class': 'rotate_then_align',
                'verification_description':
                    'Middle rectangle rotated 90° counter-clockwise, then '
                    'all three rectangles right-aligned',
                'instruction':
                    "Rotate the middle rectangle 90° counter-clockwise, then "
                    "align all three rectangles by their right edges.",
            },
        },
    },

    # ── Group Lifecycle (3 variants) ──────────────────────────────────
    #
    # Genuinely coupled: the flip / rotate operation is applied to the
    # group's bounding box, which mirrors / rotates the individual shapes'
    # positions relative to that box. Without the grouping step, the same
    # flip/rotate on individual shapes only transforms their own centers
    # and produces a different layout.
    {
        'group': 'LayoutCoupledGroupLifecycle',
        'target_object_type': 'multi_shape',
        'behavior': 'object_effect',
        'verification_method': 'composition_dispatch',
        'doc_in_name': 'doc_in_three_shapes',
        'variants': {
            'variant_flip_horizontal': {
                'steps': [
                    {'operation': 'group_all', 'params': {}},
                    {'operation': 'flip_group',
                     'params': {'direction': 0}},  # msoFlipHorizontal
                    {'operation': 'ungroup', 'params': {}},
                ],
                'interaction_class': 'group_lifecycle_flip',
                'verification_description':
                    'All three rectangles grouped, flipped horizontally, '
                    'then ungrouped',
                'instruction':
                    "Group all three rectangles together into a single "
                    "grouped shape, flip the group horizontally, then "
                    "ungroup the grouped rectangles into individual shapes.",
            },
            'variant_flip_vertical': {
                'steps': [
                    {'operation': 'group_all', 'params': {}},
                    {'operation': 'flip_group',
                     'params': {'direction': 1}},  # msoFlipVertical
                    {'operation': 'ungroup', 'params': {}},
                ],
                'interaction_class': 'group_lifecycle_flip',
                'verification_description':
                    'All three rectangles grouped, flipped vertically, '
                    'then ungrouped',
                'instruction':
                    "Group all three rectangles together into a single "
                    "grouped shape, flip the group vertically, then "
                    "ungroup the grouped rectangles into individual shapes.",
            },
            'variant_rotate': {
                'steps': [
                    {'operation': 'group_all', 'params': {}},
                    {'operation': 'rotate_group',
                     'params': {'angle': 90}},
                    {'operation': 'ungroup', 'params': {}},
                ],
                'interaction_class': 'group_lifecycle_rotate',
                'verification_description':
                    'All three rectangles grouped, rotated 90° clockwise, '
                    'then ungrouped',
                'instruction':
                    "Group all three rectangles together into a single "
                    "grouped shape, rotate the group 90° clockwise, then "
                    "ungroup the grouped rectangles into individual shapes.",
            },
        },
    },

    # ── Wrap + Position (2 variants) ──────────────────────────────────
    #
    # Genuinely coupled: a cross shape that is inline (the default wrap
    # mode) ignores explicit .Left/.Top — an inline shape's position is
    # determined by text flow. Changing the wrap type from inline to
    # square/tight establishes the "floating shape" layout context in
    # which explicit margin-relative positioning becomes meaningful.
    {
        'group': 'LayoutCoupledWrapPosition',
        'target_object_type': 'shape',
        'behavior': 'object_effect',
        'verification_method': 'composition_dispatch',
        'doc_in_name': 'doc_in_cross_with_text',
        'variants': {
            'variant_square_top_center': {
                'steps': [
                    {'operation': 'set_wrap',
                     'params': {'wrap_type': 0}},  # wdWrapSquare
                    {'operation': 'set_position',
                     'params': {
                         'h_position': -999995,  # wdShapeCenter
                         'v_position': -999999,  # wdShapeTop
                     }},
                ],
                'interaction_class': 'wrap_then_position',
                'verification_description':
                    'Cross shape wrap set to Square, then positioned at '
                    'top-center using the Position preset (within default '
                    'page margins)',
                'instruction':
                    "Apply square wrapping to the cross shape (text wraps "
                    "around the cross's rectangular bounding box), then "
                    "position it at the top-center of the page, using the "
                    "Position preset (within default page margins).",
            },
            'variant_tight_top_left': {
                'steps': [
                    {'operation': 'set_wrap',
                     'params': {'wrap_type': 1}},  # wdWrapTight
                    {'operation': 'set_position',
                     'params': {
                         'h_position': -999998,  # wdShapeLeft
                         'v_position': -999999,  # wdShapeTop
                     }},
                ],
                'interaction_class': 'wrap_then_position',
                'verification_description':
                    'Cross shape wrap set to Tight, then positioned at '
                    'top-left using the Position preset (within default '
                    'page margins)',
                'instruction':
                    "Apply tight wrapping to the cross shape (text follows "
                    "closely along the cross's edges), then position it at "
                    "the top-left corner of the page, using the Position "
                    "preset (within default page margins).",
            },
        },
    },

    # ── Position preset + Align-to-multi-reference (2 variants, new / framing-aligned)
    #
    # Genuinely coupled: the multi-shape align step's reference point is
    # the rightmost (or topmost) shape among the selection. Repositioning
    # one shape first via a standard Position gallery preset can change
    # which shape is that reference, or can change the reference value
    # itself, altering where the other shapes land.
    #
    # Uses existing atomic Ribbon/gallery operations rather than exact
    # coordinate movement.
    {
        'group': 'LayoutCoupledPositionAlignMulti',
        'target_object_type': 'multi_shape',
        'behavior': 'object_effect',
        'verification_method': 'composition_dispatch',
        'doc_in_name': 'doc_in_position_align_shapes',
        'needs_position_align_shapes': True,
        'variants': {
            'variant_bottom_right_align_right': {
                'steps': [
                    {'operation': 'set_position',
                     'params': {
                         'shape_index': 1,
                         'h_position': -999996,  # wdShapeRight
                         'v_position': -999997,  # wdShapeBottom
                     }},
                    {'operation': 'align_shapes',
                     'params': {'align_type': 2}},  # msoAlignRights
                ],
                'interaction_class': 'position_then_align_multi',
                'verification_description':
                    'First rectangle positioned at bottom-right using the '
                    'Position preset, then all three rectangles right-aligned '
                    'to the rightmost edge among them',
                'instruction':
                    "Position the first rectangle at the bottom-right corner "
                    "of the page using the Position preset (within default "
                    "page margins), then align all three rectangles' right "
                    "edges to the rightmost edge among them.",
            },
            'variant_top_center_align_top': {
                'steps': [
                    {'operation': 'set_position',
                     'params': {
                         'shape_index': 1,
                         'h_position': -999995,  # wdShapeCenter
                         'v_position': -999999,  # wdShapeTop
                     }},
                    {'operation': 'align_shapes',
                     'params': {'align_type': 3}},  # msoAlignTops
                ],
                'interaction_class': 'position_then_align_multi',
                'verification_description':
                    'First rectangle positioned at top-center using the '
                    'Position preset, then all three rectangles top-aligned '
                    'to the topmost edge among them',
                'instruction':
                    "Position the first rectangle at the top-center of the "
                    "page using the Position preset (within default page "
                    "margins), then align all three rectangles' top edges "
                    "to the topmost edge among them.",
            },
        },
    },

    # ── Multi-Step Layout-Coupled Chains (2 variants) ─────────────────
    #
    # Both chains have at least one genuine coupling edge (the rotate→align
    # or group→flip step), confirmed empirically. The trailing distribute
    # step is order-agnostic on its own, but the chain as a whole is
    # order-sensitive because of the preceding coupled step.
    {
        'group': 'LayoutCoupledMultiStep',
        'target_object_type': 'multi_shape',
        'behavior': 'object_effect',
        'verification_method': 'composition_dispatch',
        'variants': {
            'variant_rotate_align_distribute': {
                'steps': [
                    {'operation': 'rotate_shape',
                     'params': {'shape_index': 2, 'angle': 90}},
                    {'operation': 'align_shapes',
                     'params': {'align_type': 3}},  # msoAlignTops
                    {'operation': 'distribute_shapes',
                     'params': {'distribute_type': 0}},  # msoDistributeHorizontally
                ],
                'interaction_class': 'rotate_align_distribute',
                'verification_description':
                    'Middle rectangle rotated 90° clockwise, all three '
                    'rectangles top-aligned, then distributed horizontally '
                    'with equal spacing',
                'instruction':
                    "Rotate the middle rectangle 90° clockwise, align all "
                    "three rectangles by their top edges, then distribute "
                    "all three rectangles evenly horizontally with equal "
                    "spacing between them.",
                'doc_in_name': 'doc_in_three_shapes_tall',
                'needs_tall_shapes': True,
            },
            'variant_group_flip_ungroup_distribute': {
                'steps': [
                    {'operation': 'group_all', 'params': {}},
                    {'operation': 'flip_group',
                     'params': {'direction': 0}},  # msoFlipHorizontal
                    {'operation': 'ungroup', 'params': {}},
                    {'operation': 'distribute_shapes',
                     'params': {'distribute_type': 0}},  # msoDistributeHorizontally
                ],
                'interaction_class': 'group_flip_ungroup_distribute',
                'verification_description':
                    'All three rectangles grouped, flipped horizontally, '
                    'ungrouped, then distributed horizontally with equal '
                    'spacing',
                'instruction':
                    "Group all three rectangles together into a single "
                    "grouped shape, flip the group horizontally, ungroup "
                    "the grouped rectangles into individual shapes, then "
                    "distribute all three rectangles evenly horizontally "
                    "with equal spacing between them.",
                'doc_in_name': 'doc_in_three_shapes_uneven',
            },
        },
    },
]


# ═════════════════════════════════════════════════════════════════════════
# Generator class
# ═════════════════════════════════════════════════════════════════════════

class LayoutCoupledGenerator:
    """Generator for layout-coupled interaction test cases."""

    def __init__(self, output_dir: Path = None, hidden: bool = False):
        self.output_dir = output_dir or (BENCHMARK_OUTPUT_DIR / TEST_SET)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.inputs_dir = self.output_dir / 'inputs'
        self.inputs_dir.mkdir(parents=True, exist_ok=True)
        self.hidden = hidden

    def generate_all(self, command: str = None):
        """Generate all layout-coupled test cases."""
        wd = ensure_word(visible=not self.hidden)
        results = {}

        for group_def in CASE_GROUPS:
            group_name = group_def['group']
            if command and command != group_name:
                continue
            for variant_name, variant_def in group_def['variants'].items():
                case_key = f'{group_name}/{variant_name}'
                # Merge group-level and variant-level fields
                case_def = self._merge_case_def(group_def, variant_def,
                                                variant_name)
                try:
                    ok = self._generate_case(wd, group_name, variant_name,
                                             case_def)
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
            variants_instr = {}
            for variant_name, variant_def in group_def['variants'].items():
                variants_instr[variant_name] = variant_def['instruction']
            commands[group_name] = variants_instr
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
            'target_object_type': (variant_def.get('target_object_type')
                                   or group_def.get('target_object_type')),
            'behavior': (variant_def.get('behavior')
                         or group_def.get('behavior')),
            'verification_method': (variant_def.get('verification_method')
                                    or group_def.get('verification_method')),
            'steps': variant_def['steps'],
            'interaction_class': variant_def['interaction_class'],
            'verification_description': variant_def['verification_description'],
            'instruction': variant_def['instruction'],
        }
        # Optional fields from variant or group
        for key in ('doc_in_name', 'needs_tall_shapes',
                    'needs_spread_tall_shapes',
                    'needs_position_align_shapes'):
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

        # Determine doc_in
        doc_in_ref, doc_in_path = self._resolve_doc_in(wd, case_def,
                                                        variant_dir)

        # Generate gold via COM
        self._generate_gold(wd, case_def, doc_in_path, gold_path)

        # Build and write spec
        atomic_intents = build_layout_coupled_atomic_intents(case_def['steps'])
        spec = build_spec(
            behavior=case_def['behavior'],
            composition_type='layout_coupled',
            interaction_class=case_def['interaction_class'],
            target_object_type=case_def['target_object_type'],
            steps=case_def['steps'],
            precondition=self._build_precondition(case_def),
            verification_method=case_def['verification_method'],
            verification_description=case_def['verification_description'],
            doc_in_ref=doc_in_ref,
            variant_name=variant_name,
            atomic_intents=atomic_intents,
        )
        write_spec(spec_path, spec)
        return True

    # ─────────────────────────────────────────────────────────────────
    # Doc_in resolution
    # ─────────────────────────────────────────────────────────────────

    def _resolve_doc_in(self, wd, case_def, variant_dir):
        """Return (relative_ref, absolute_path) for the doc_in."""
        doc_in_name = case_def['doc_in_name']

        if case_def.get('needs_spread_tall_shapes'):
            dest = self.inputs_dir / f'{doc_in_name}.docx'
            if not dest.exists():
                create_spread_tall_shapes_doc(wd, dest)
                print(f'  Created: {doc_in_name}.docx')
            return f'../../inputs/{doc_in_name}.docx', dest
        if case_def.get('needs_tall_shapes'):
            dest = self.inputs_dir / f'{doc_in_name}.docx'
            if not dest.exists():
                create_tall_shapes_doc(wd, dest)
                print(f'  Created: {doc_in_name}.docx')
            return f'../../inputs/{doc_in_name}.docx', dest
        if case_def.get('needs_position_align_shapes'):
            dest = self.inputs_dir / f'{doc_in_name}.docx'
            if not dest.exists():
                self._create_position_align_shapes_doc(wd, dest)
                print(f'  Created: {doc_in_name}.docx')
            return f'../../inputs/{doc_in_name}.docx', dest

        # All other inputs are shared from existing test sets
        path = get_shared_input(self.inputs_dir, doc_in_name)
        return f'../../inputs/{doc_in_name}.docx', path

    # ─────────────────────────────────────────────────────────────────
    # Local doc_in creators for framing-aligned groups
    # ─────────────────────────────────────────────────────────────────

    def _create_position_align_shapes_doc(self, wd, dest_path):
        """Create a document with three 80x80 rectangles, all anchored to
        the page margin (not the default paragraph). Shape 1 is the
        leftmost / topmost initially, so applying a Position preset before
        multi-shape alignment changes the alignment reference context.
        """
        doc = wd.Documents.Add()
        time.sleep(0.3)
        positions = [(20, 20), (120, 120), (220, 220)]
        for left, top in positions:
            shape = doc.Shapes.AddShape(1, left, top, 80, 80)
            try:
                # Anchor to margin so .Left / .Top are predictable
                shape.RelativeHorizontalPosition = 0  # wdRelativeHorizontalPositionMargin
                shape.RelativeVerticalPosition = 0    # wdRelativeVerticalPositionMargin
                shape.Left = left
                shape.Top = top
            except Exception:
                pass
            time.sleep(0.2)
        doc.SaveAs(str(Path(dest_path).absolute()), FileFormat=16)
        doc.Close(False)
        time.sleep(0.3)

    # ─────────────────────────────────────────────────────────────────
    # Gold generation — dispatches by interaction_class
    # ─────────────────────────────────────────────────────────────────

    def _generate_gold(self, wd, case_def, doc_in_path, gold_path):
        """Open doc_in, execute all steps via COM, save as gold."""
        ic = case_def['interaction_class']

        if ic == 'rotate_then_align':
            self._gold_rotate_align(wd, case_def, doc_in_path, gold_path)
        elif ic in ('group_lifecycle_flip', 'group_lifecycle_rotate'):
            self._gold_group_lifecycle(wd, case_def, doc_in_path, gold_path)
        elif ic == 'wrap_then_position':
            self._gold_wrap_position(wd, case_def, doc_in_path, gold_path)
        elif ic == 'position_then_align_multi':
            self._gold_position_align_multi(wd, case_def, doc_in_path, gold_path)
        elif ic == 'rotate_align_distribute':
            self._gold_rotate_align_distribute(wd, case_def, doc_in_path,
                                               gold_path)
        elif ic == 'group_flip_ungroup_distribute':
            self._gold_group_flip_ungroup_distribute(wd, case_def, doc_in_path,
                                                     gold_path)
        else:
            raise ValueError(
                f'Unknown interaction_class: {ic}')

    # ─────────────────────────────────────────────────────────────────
    # Shape selection helper
    # ─────────────────────────────────────────────────────────────────

    def _select_all_shapes(self, doc, wd):
        """Select all shapes in the document (multi-select)."""
        doc.Shapes(1).Select()
        for i in range(2, doc.Shapes.Count + 1):
            doc.Shapes(i).Select(False)
        time.sleep(0.2)

    # ─────────────────────────────────────────────────────────────────
    # Type: Rotate + Align (visual-bbox-aware, genuinely coupled)
    # ─────────────────────────────────────────────────────────────────

    def _gold_rotate_align(self, wd, case_def, doc_in_path, gold_path):
        """Rotate one shape, then align all shapes.

        Uses ShapeRange.Align (not ExecuteMso) because the former uses
        the visual (rotated) bounding box — which is what makes this
        composition layout-coupled.
        """
        doc = open_doc(wd, doc_in_path)
        try:
            steps = case_def['steps']

            # Step A: Rotate specified shape
            shape_idx = steps[0]['params']['shape_index']
            angle = steps[0]['params']['angle']
            doc.Shapes(shape_idx).IncrementRotation(angle)
            time.sleep(0.3)

            # Step B: Align all shapes via ShapeRange (visual bbox aware)
            step_b = steps[1]
            shape_range = doc.Shapes.Range(
                list(range(1, doc.Shapes.Count + 1))
            )
            shape_range.Align(step_b['params']['align_type'], False)
            time.sleep(0.5)

            save_as_gold(doc, gold_path)
        finally:
            close_doc(doc)

    # ─────────────────────────────────────────────────────────────────
    # Type 3: Group Lifecycle
    # ─────────────────────────────────────────────────────────────────

    def _gold_group_lifecycle(self, wd, case_def, doc_in_path, gold_path):
        """Group shapes, apply operation on group, ungroup."""
        doc = open_doc(wd, doc_in_path)
        try:
            steps = case_def['steps']

            # Step A: Group all shapes
            shape_count = doc.Shapes.Count
            shape_range = doc.Shapes.Range(
                list(range(1, shape_count + 1))
            )
            group = shape_range.Group()
            time.sleep(0.3)

            # Step B: Apply operation on group
            step_b = steps[1]
            if step_b['operation'] == 'flip_group':
                group.Flip(step_b['params']['direction'])
                time.sleep(0.3)
            elif step_b['operation'] == 'rotate_group':
                group.IncrementRotation(step_b['params']['angle'])
                time.sleep(0.3)

            # Step C: Ungroup
            group.Ungroup()
            time.sleep(0.3)

            save_as_gold(doc, gold_path)
        finally:
            close_doc(doc)

    # ─────────────────────────────────────────────────────────────────
    # Type 4: Wrap + Position
    # ─────────────────────────────────────────────────────────────────

    def _gold_wrap_position(self, wd, case_def, doc_in_path, gold_path):
        """Set text wrapping, then position the shape."""
        doc = open_doc(wd, doc_in_path)
        try:
            steps = case_def['steps']
            shape = doc.Shapes(1)

            # Step A: Set text wrapping
            wrap_type = steps[0]['params']['wrap_type']
            shape.WrapFormat.Type = wrap_type
            time.sleep(0.3)

            # Step B: Set position relative to page margin
            h_pos = steps[1]['params']['h_position']
            v_pos = steps[1]['params']['v_position']
            shape.RelativeHorizontalPosition = 0  # wdRelativeHorizontalPositionMargin
            shape.RelativeVerticalPosition = 0    # wdRelativeVerticalPositionMargin
            shape.Left = h_pos
            shape.Top = v_pos
            time.sleep(0.3)

            save_as_gold(doc, gold_path)
        finally:
            close_doc(doc)

    # ─────────────────────────────────────────────────────────────────
    # Type: Position preset + Multi-align (reference-based coupling)
    # ─────────────────────────────────────────────────────────────────

    def _gold_position_align_multi(self, wd, case_def, doc_in_path, gold_path):
        """Apply a standard Position preset to one shape, then align all
        shapes by the given edge.

        The alignment step's reference point (rightmost/topmost) depends
        on where all shapes currently are. Repositioning one shape first
        with the same gallery command covered by atomic bench changes
        that context without introducing an exact-coordinate move helper.
        """
        doc = open_doc(wd, doc_in_path)
        try:
            steps = case_def['steps']

            # Step A: Apply position preset to specified shape
            pos_params = steps[0]['params']
            shape = doc.Shapes(pos_params.get('shape_index', 1))
            shape.RelativeHorizontalPosition = 0  # margin
            shape.RelativeVerticalPosition = 0    # margin
            shape.Left = pos_params['h_position']
            shape.Top = pos_params['v_position']
            time.sleep(0.3)

            # Step B: Align all shapes
            shape_range = doc.Shapes.Range(
                list(range(1, doc.Shapes.Count + 1))
            )
            shape_range.Align(steps[1]['params']['align_type'], False)
            time.sleep(0.5)

            save_as_gold(doc, gold_path)
        finally:
            close_doc(doc)

    # ─────────────────────────────────────────────────────────────────
    # Multi-Step: Rotate -> Align -> Distribute
    # ─────────────────────────────────────────────────────────────────

    def _gold_rotate_align_distribute(self, wd, case_def, doc_in_path,
                                      gold_path):
        """Rotate one shape, align all, then distribute all.

        Uses ShapeRange.Align/Distribute (not ExecuteMso) to correctly
        handle visual bounding boxes of rotated shapes.
        """
        doc = open_doc(wd, doc_in_path)
        try:
            steps = case_def['steps']

            # Step A: Rotate specified shape
            shape_idx = steps[0]['params']['shape_index']
            angle = steps[0]['params']['angle']
            doc.Shapes(shape_idx).IncrementRotation(angle)
            time.sleep(0.3)

            shape_range = doc.Shapes.Range(
                list(range(1, doc.Shapes.Count + 1))
            )

            # Step B: Align all shapes
            shape_range.Align(steps[1]['params']['align_type'], False)
            time.sleep(0.5)

            # Step C: Distribute all shapes
            shape_range.Distribute(
                steps[2]['params']['distribute_type'], False)
            time.sleep(0.5)

            save_as_gold(doc, gold_path)
        finally:
            close_doc(doc)

    # ─────────────────────────────────────────────────────────────────
    # Multi-Step: Group -> Rotate -> Ungroup -> Align
    # ─────────────────────────────────────────────────────────────────

    def _gold_group_flip_ungroup_distribute(self, wd, case_def, doc_in_path,
                                             gold_path):
        """Group all, flip group, ungroup, then distribute individual shapes."""
        doc = open_doc(wd, doc_in_path)
        try:
            steps = case_def['steps']

            # Step A: Group all shapes
            shape_count = doc.Shapes.Count
            shape_range = doc.Shapes.Range(
                list(range(1, shape_count + 1))
            )
            group = shape_range.Group()
            time.sleep(0.3)

            # Step B: Flip the group
            group.Flip(steps[1]['params']['direction'])
            time.sleep(0.3)

            # Step C: Ungroup
            group.Ungroup()
            time.sleep(0.3)

            # Step D: Distribute all shapes
            sr = doc.Shapes.Range(
                list(range(1, doc.Shapes.Count + 1))
            )
            sr.Distribute(steps[3]['params']['distribute_type'], False)
            time.sleep(0.5)

            save_as_gold(doc, gold_path)
        finally:
            close_doc(doc)

    # ─────────────────────────────────────────────────────────────────
    # Helpers
    # ─────────────────────────────────────────────────────────────────

    def _build_precondition(self, case_def):
        """Build the precondition dict for the spec."""
        doc_in_name = case_def['doc_in_name']

        if doc_in_name == 'doc_in_three_shapes':
            return {
                'type': 'three_shapes',
                'description':
                    'Document with three 100x100 rectangle shapes at '
                    '(50,50), (100,100), (150,150)',
            }
        if doc_in_name == 'doc_in_three_shapes_uneven':
            return {
                'type': 'three_shapes_uneven',
                'description':
                    'Document with three 100x100 rectangles at (50,50), '
                    '(70,70), (150,150) — overlapping and unevenly spaced',
            }
        if doc_in_name == 'doc_in_three_shapes_tall':
            return {
                'type': 'three_shapes_tall',
                'description':
                    'Document with three tall narrow rectangles (60x120) at '
                    '(50,50), (150,100), (250,150)',
            }
        if doc_in_name == 'doc_in_three_shapes_tall_spread':
            return {
                'type': 'three_shapes_tall_spread',
                'description':
                    'Document with three tall narrow rectangles (60x120) '
                    'unevenly spread at (50,50), (170,200), (400,550)',
            }
        if doc_in_name == 'doc_in_cross_with_text':
            return {
                'type': 'shape_with_text',
                'description':
                    'Document with a cross shape and surrounding text',
            }
        if doc_in_name == 'doc_in_position_align_shapes':
            return {
                'type': 'three_shapes_stagger',
                'description':
                    'Document with three 80x80 rectangles at (20,20), '
                    '(120,120), (220,220). Shape 1 is initially the '
                    'leftmost and topmost so a Position preset can change '
                    'the align reference context.',
            }

        return {
            'type': 'unknown',
            'description': f'Input document: {doc_in_name}',
        }
