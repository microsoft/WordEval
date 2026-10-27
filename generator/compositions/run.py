"""
Entry point for composition test case generation.

Usage:
    python -m generator.compositions.run --test-set interaction_composition_x_prerequisite_chain
    python -m generator.compositions.run --test-set all
    python -m generator.compositions.run --test-set all --instructions-only
"""

import sys
import argparse
from pathlib import Path

GENERATORS = {
    'additive_composition_x_content_formatting': 'additive.content_formatting',
    'additive_composition_x_object_effect': 'additive.object_effect',
    'additive_composition_x_content_operations': 'additive.content_operations',
    'interaction_composition_x_prerequisite_chain': 'interaction.prerequisite',
    'interaction_composition_x_scope_changing': 'interaction.scope_changing',
    'interaction_composition_x_layout_coupled': 'interaction.layout_coupled',
}


def _write_instructions(test_set: str):
    """Write instructions JSON for a test set (no COM/Word needed)."""
    from generator.compositions.interaction.shared import write_instructions_json

    if test_set == 'additive_composition_x_content_formatting':
        from generator.compositions.additive.content_formatting import (
            FONT_VARIANTS, PARAGRAPH_VARIANTS, LIST_VARIANTS,
        )
        commands = {}
        for cmd, v in [('AdditiveFontComposition', FONT_VARIANTS),
                       ('AdditiveParagraphComposition', PARAGRAPH_VARIANTS),
                       ('AdditiveListComposition', LIST_VARIANTS)]:
            commands[cmd] = {vn: vd['instruction'] for vn, vd in v.items()}
        write_instructions_json(test_set, commands)

    elif test_set == 'additive_composition_x_object_effect':
        from generator.compositions.additive.object_effect import (
            _write_all_instructions,
        )
        _write_all_instructions()

    elif test_set == 'additive_composition_x_content_operations':
        from generator.compositions.additive.content_operations import (
            DOCUMENT_ENRICHMENT_VARIANTS, DOCUMENT_CLEANUP_VARIANTS,
        )
        commands = {}
        for cmd, v in [
            ('AdditiveDocumentEnrichmentComposition',
             DOCUMENT_ENRICHMENT_VARIANTS),
            ('AdditiveDocumentCleanupComposition',
             DOCUMENT_CLEANUP_VARIANTS),
        ]:
            commands[cmd] = {vn: vd['instruction'] for vn, vd in v.items()}
        write_instructions_json(test_set, commands)

    elif test_set == 'interaction_composition_x_prerequisite_chain':
        from generator.compositions.interaction.prerequisite import (
            CASE_GROUPS,
        )
        commands = {}
        for group_def in CASE_GROUPS:
            group_name = group_def['group']
            commands[group_name] = {
                vn: vd['instruction']
                for vn, vd in group_def['variants'].items()
            }
        write_instructions_json(test_set, commands)

    elif test_set == 'interaction_composition_x_scope_changing':
        from generator.compositions.interaction.scope_changing import (
            CASE_GROUPS,
        )
        commands = {}
        for group_def in CASE_GROUPS:
            group_name = group_def['group']
            commands[group_name] = {
                vn: vd['instruction']
                for vn, vd in group_def['variants'].items()
            }
        write_instructions_json(test_set, commands)

    elif test_set == 'interaction_composition_x_layout_coupled':
        from generator.compositions.interaction.layout_coupled import (
            CASE_GROUPS,
        )
        commands = {}
        for group_def in CASE_GROUPS:
            group_name = group_def['group']
            commands[group_name] = {
                vn: vd['instruction']
                for vn, vd in group_def['variants'].items()
            }
        write_instructions_json(test_set, commands)

    print(f"  Wrote instructions/{test_set}.json")


def main():
    parser = argparse.ArgumentParser(
        description='Generate composition test cases',
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        '--test-set', '-s',
        choices=list(GENERATORS.keys()) + ['all'],
        required=True,
        help='Test set to generate (or "all")',
    )
    parser.add_argument('--command', '-c', help='Generate specific case only')
    parser.add_argument('--output', '-o', help='Output directory')
    parser.add_argument('--hidden', action='store_true', help='Run Word hidden')
    parser.add_argument(
        '--instructions-only', action='store_true',
        help='Only write instructions JSON (no COM/Word needed)',
    )

    args = parser.parse_args()

    test_sets = list(GENERATORS.keys()) if args.test_set == 'all' else [args.test_set]

    if args.instructions_only:
        for ts in test_sets:
            _write_instructions(ts)
        return

    for ts in test_sets:
        print(f"\n{'='*60}")
        print(f"Generating: {ts}")
        print(f"{'='*60}")

        if ts == 'additive_composition_x_content_formatting':
            from generator.compositions.additive.content_formatting import (
                AdditiveContentFormattingGenerator,
            )
            output_dir = Path(args.output) if args.output else None
            gen = AdditiveContentFormattingGenerator(
                output_dir=output_dir, hidden=args.hidden
            )
            gen.generate_all(command=args.command)

        elif ts == 'additive_composition_x_object_effect':
            from generator.compositions.additive.object_effect import (
                AdditiveObjectEffectGenerator,
            )
            output_dir = Path(args.output) if args.output else None
            gen = AdditiveObjectEffectGenerator(
                output_dir=output_dir, hidden=args.hidden
            )
            gen.generate_all(command=args.command)

        elif ts == 'additive_composition_x_content_operations':
            from generator.compositions.additive.content_operations import (
                AdditiveContentOperationsGenerator,
            )
            output_dir = Path(args.output) if args.output else None
            gen = AdditiveContentOperationsGenerator(
                output_dir=output_dir, hidden=args.hidden
            )
            gen.generate_all(command=args.command)

        elif ts == 'interaction_composition_x_prerequisite_chain':
            from generator.compositions.interaction.prerequisite import (
                PrerequisiteGenerator,
            )
            output_dir = Path(args.output) if args.output else None
            gen = PrerequisiteGenerator(
                output_dir=output_dir, hidden=args.hidden
            )
            gen.generate_all(command=args.command)

        elif ts == 'interaction_composition_x_scope_changing':
            from generator.compositions.interaction.scope_changing import (
                ScopeChangingGenerator,
            )
            output_dir = Path(args.output) if args.output else None
            gen = ScopeChangingGenerator(
                output_dir=output_dir, hidden=args.hidden
            )
            gen.generate_all(command=args.command)

        elif ts == 'interaction_composition_x_layout_coupled':
            from generator.compositions.interaction.layout_coupled import (
                LayoutCoupledGenerator,
            )
            output_dir = Path(args.output) if args.output else None
            gen = LayoutCoupledGenerator(
                output_dir=output_dir, hidden=args.hidden
            )
            gen.generate_all(command=args.command)


if __name__ == '__main__':
    main()
