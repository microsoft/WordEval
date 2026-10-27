#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Benchmark runner script - ONE COMMAND TO RULE THEM ALL

Two modes:

1. GENERATE mode: Generate test cases for a specific test set
   python generator/test_sets/run.py --test-set content_formatting_x_core_actionable --limit 5
   python generator/test_sets/run.py --test-set content_formatting_x_core_actionable --command Bold
   python generator/test_sets/run.py --test-set object_effect_x_core_actionable

2. VERIFY mode: Verify existing benchmark and generate final_summary.txt
   python generator/test_sets/run.py --verify --benchmark-dir ./benchmark/object_effect_x_core_actionable

   # Short form
   python generator/test_sets/run.py -v -b ./benchmark/object_effect_x_core_actionable
"""

import sys
from importlib import import_module
from pathlib import Path

# Add project root to Python path
project_root = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(project_root))


# Available generators mapped by test set name (BEHAVIOR × SCOPE)
GENERATORS = {
    'content_formatting_x_core_actionable': 'content_formatting_core_actionable',
    'content_formatting_x_core_parameterized': 'content_formatting_core_parameterized',
    'content_operations_x_core_actionable': 'content_operations_x_core_actionable',
    'content_operations_x_core_dialog': 'content_operations_x_core_dialog',
    'content_operations_x_core_parameterized': 'content_operations_core_parameterized',
    'content_operations_x_not_in_ribbon': 'content_operations_x_not_in_ribbon',
    'content_operations_x_specialized_actionable': 'content_operations_specialized_actionable',
    'content_operations_x_contextual_actionable': 'content_operations_x_contextual_actionable',
    'object_effect_x_core_actionable': 'object_effect_core_actionable',
    'object_effect_x_core_parameterized': 'object_effect_core_parameterized',
    'object_effect_x_contextual_actionable': 'object_effect_contextual_actionable',
    'object_effect_x_contextual_parameterized': 'object_effect_contextual_parameterized',
    # Composition test sets (generator/compositions/)
    'additive_composition_x_content_formatting': 'compositions.additive.content_formatting',
    'additive_composition_x_object_effect': 'compositions.additive.object_effect',
    'additive_composition_x_content_operations': 'compositions.additive.content_operations',
    'interaction_composition_x_prerequisite_chain': 'compositions.interaction.prerequisite',
    'interaction_composition_x_scope_changing': 'compositions.interaction.scope_changing',
    'interaction_composition_x_layout_coupled': 'compositions.interaction.layout_coupled',
}

COMPOSITION_TEST_SETS = {
    test_set for test_set, module in GENERATORS.items()
    if module.startswith('compositions.')
}

TRACK_TEST_SETS = {
    'atomic': [
        test_set for test_set in GENERATORS
        if test_set not in COMPOSITION_TEST_SETS
    ],
    'compositional': [
        test_set for test_set in GENERATORS
        if test_set in COMPOSITION_TEST_SETS
    ],
}
TRACK_TEST_SETS['all'] = TRACK_TEST_SETS['atomic'] + TRACK_TEST_SETS['compositional']

HIDDEN_SUPPORTED_GENERATORS = {
    'content_formatting_core_parameterized',
    'content_operations_x_core_dialog',
    'content_operations_x_not_in_ribbon',
}


def benchmark_dir_for(test_set: str, benchmark_root: str = None) -> str:
    root = Path(benchmark_root or 'benchmark')
    return str(root / test_set)


def run_verify(benchmark_dir: str) -> None:
    sys.argv = ['summarize_test_cases.py', '--benchmark-dir', benchmark_dir]
    from generator.test_sets.summarize_test_cases import main
    main()


def build_generate_args(args, test_set: str, multi_set: bool = False) -> tuple[str, list[str]]:
    generator_module = GENERATORS[test_set]

    if test_set in COMPOSITION_TEST_SETS:
        gen_args = ['compositions.run.py', '--test-set', test_set]
        if args.command:
            gen_args.extend(['--command', args.command])
        if args.output:
            output = Path(args.output) / test_set if multi_set else Path(args.output)
            gen_args.extend(['--output', str(output)])
        if args.hidden:
            gen_args.append('--hidden')
        if args.limit:
            print(f"[WARN] --limit is not supported by {test_set}; ignoring it.")
        if args.timeout_gen or args.timeout != 20.0:
            print(f"[WARN] --timeout is not supported by {test_set}; ignoring it.")
        return 'generator.compositions.run', gen_args

    gen_args = [f'{generator_module}.py']
    if args.command:
        gen_args.extend(['--command', args.command])
    if args.limit:
        gen_args.extend(['--limit', str(args.limit)])
    if args.output:
        output = Path(args.output) / test_set if multi_set else Path(args.output)
        gen_args.extend(['--output', str(output)])
    if args.timeout_gen:
        gen_args.extend(['--timeout-gen', str(args.timeout_gen)])
    elif args.timeout != 20.0:
        gen_args.extend(['--timeout-gen', str(args.timeout)])
    if args.hidden and generator_module in HIDDEN_SUPPORTED_GENERATORS:
        gen_args.append('--hidden')
    elif args.hidden:
        print(
            f"[WARN] --hidden is not supported by {test_set}; "
            "running with the generator default Word visibility."
        )
    return f'generator.test_sets.{generator_module}', gen_args


def run_generate(args, test_set: str, multi_set: bool = False) -> None:
    module_name, gen_args = build_generate_args(args, test_set, multi_set=multi_set)
    sys.argv = gen_args
    module = import_module(module_name)
    module.main()


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(
        description='Benchmark generator and verifier - ONE COMMAND TO RULE THEM ALL',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:

GENERATE mode:
  python generator/test_sets/run.py --test-set content_formatting_x_core_actionable --limit 5
  python generator/test_sets/run.py --test-set content_formatting_x_core_actionable --command Bold
  python generator/test_sets/run.py --test-set object_effect_x_core_actionable

VERIFY mode:
  python generator/test_sets/run.py --verify --benchmark-dir ./benchmark/object_effect_x_core_actionable
  python generator/test_sets/run.py --verify --track atomic

  # Short form
  python generator/test_sets/run.py -v -b ./benchmark/object_effect_x_core_actionable
        """
    )

    # Mode selection
    parser.add_argument(
        '--verify', '-v',
        action='store_true',
        help='Run verification mode (verify doc_in vs gold and generate summary)'
    )

    # Verify mode arguments
    parser.add_argument(
        '--benchmark-dir', '-b',
        help='Benchmark directory (required for --verify mode)'
    )
    parser.add_argument(
        '--benchmark-root',
        default='benchmark',
        help='Benchmark root for --verify --track (default: ./benchmark)'
    )

    # Generate mode arguments
    parser.add_argument(
        '--test-set', '-s',
        choices=list(GENERATORS.keys()),
        required=False,
        help='Test set name (BEHAVIOR × SCOPE) to generate tests for'
    )
    parser.add_argument(
        '--command', '-c',
        help='Generate for specific command (generate mode only)'
    )
    parser.add_argument(
        '--track',
        choices=['atomic', 'compositional', 'all'],
        help='Generate or verify all test sets for a track sequentially'
    )
    parser.add_argument(
        '--limit', '-l',
        type=int,
        help='Limit number of commands to generate (generate mode only)'
    )
    parser.add_argument(
        '--output', '-o',
        help='Output directory (generate mode only)'
    )
    parser.add_argument(
        '--timeout', '-t',
        type=float,
        default=20.0,
        help='Timeout in seconds per test case (generate mode only, default: 20.0)'
    )
    parser.add_argument(
        '--timeout-gen',
        type=float,
        help='Timeout for generation in seconds (generate mode only, overrides --timeout)'
    )
    parser.add_argument(
        '--hidden',
        action='store_true',
        help='Run Word hidden when the selected generator supports it'
    )

    args = parser.parse_args()

    if args.verify:
        # SUMMARIZE MODE (legacy --verify flag for backward compatibility)
        if args.track:
            if args.benchmark_dir:
                print("[ERROR] Use either --benchmark-dir or --track for --verify, not both")
                sys.exit(1)
            for test_set in TRACK_TEST_SETS[args.track]:
                benchmark_dir = benchmark_dir_for(test_set, args.benchmark_root)
                print(f"\n{'='*60}")
                print(f"Verifying: {test_set}")
                print(f"{'='*60}")
                run_verify(benchmark_dir)
            sys.exit(0)

        if not args.benchmark_dir:
            print("[ERROR] --benchmark-dir is required for --verify mode")
            print("Usage: python run.py --verify --benchmark-dir ./benchmark/content_formatting_x_core_actionable")
            print("   or: python run.py --verify --track atomic")
            sys.exit(1)

        run_verify(args.benchmark_dir)

    else:
        # GENERATE MODE - determine test set
        if args.track:
            if args.test_set:
                print("[ERROR] Use either --test-set or --track, not both")
                sys.exit(1)
            if args.command:
                print("[ERROR] --command requires --test-set; it is ambiguous with --track")
                sys.exit(1)
            for test_set in TRACK_TEST_SETS[args.track]:
                print(f"\n{'='*60}")
                print(f"Generating: {test_set}")
                print(f"{'='*60}")
                run_generate(args, test_set, multi_set=True)
            sys.exit(0)

        if not args.test_set:
            print("[ERROR] --test-set is required for generate mode")
            print("Usage: python run.py --test-set object_effect_x_core_actionable")
            print("   or: python run.py --track atomic")
            print("\nAvailable test sets:")
            for ts in sorted(GENERATORS.keys()):
                print(f"  - {ts}")
            sys.exit(1)

        test_set = args.test_set
        run_generate(args, test_set)
