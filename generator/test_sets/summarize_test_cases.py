"""
Summarize generated benchmark test cases and collect statistics.

This script scans the generated benchmark directory and:
1. Finds all test cases (doc_in.docx + gold.docx pairs)
2. Compares doc_in vs gold to categorize test cases:
   - xml_effect_execution: Command modified the document
   - no_xml_effect_execution: Command had no XML effect (needs review)
3. Collects statistics and outputs structured summary

Usage:
    python summarize_test_cases.py --benchmark-dir ./benchmark/content_formatting_x_core_actionable
    python summarize_test_cases.py -b ./benchmark/content_formatting_x_core_actionable --output ./custom_output
"""

import os
import re
import json
import zipfile
import argparse
from pathlib import Path
from typing import Dict, List, Tuple
from datetime import datetime


# ---------- XML comparison helpers ----------

IGNORABLE_PATTERNS = re.compile(
    r'\s+w:rsid[^=]*="[^"]*"'     # rsid variants
    r'|\s+w:rsidRPr="[^"]*"'
    r'|\s+w14:paraId="[^"]*"'     # paragraph IDs (change on every save)
    r'|\s+w14:textId="[^"]*"'     # text IDs (change on every save)
    r'|\s+m:mathPr[^>]*>.*?</m:mathPr>'  # math props sometimes reorder
, re.S)


def read_main_xml(docx_path: str) -> str:
    """Read and normalize the main document XML from a .docx file."""
    with zipfile.ZipFile(docx_path, "r") as z:
        xml = z.read("word/document.xml").decode("utf-8", errors="ignore")
    # normalize: drop ignorable attrs and extra whitespace
    xml = IGNORABLE_PATTERNS.sub("", xml)
    xml = re.sub(r">\s+<", "><", xml)
    return xml


def read_footnotes_xml(docx_path: str) -> str:
    """
    Read and normalize footnotes/endnotes XML parts from a .docx file.

    Includes:
    - word/footnotes.xml
    - word/endnotes.xml

    Returns:
        Combined normalized content of footnote/endnote XML files
    """
    xml_parts = []

    with zipfile.ZipFile(docx_path, "r") as z:
        file_list = z.namelist()
        for filename in sorted(file_list):
            if filename in ("word/footnotes.xml", "word/endnotes.xml"):
                xml_parts.append(z.read(filename).decode("utf-8", errors="ignore"))

    xml = "\n".join(xml_parts)

    # normalize: drop ignorable attrs and extra whitespace
    xml = IGNORABLE_PATTERNS.sub("", xml)
    xml = re.sub(r">\s+<", "><", xml)
    return xml


def read_supplementary_xml(docx_path: str) -> str:
    """
    Read and normalize supplementary XML parts from a .docx file.

    Includes:
    - word/header*.xml (all headers)
    - word/footer*.xml (all footers)

    Returns:
        Combined normalized content of all supplementary XML files
    """
    xml_parts = []

    with zipfile.ZipFile(docx_path, "r") as z:
        file_list = z.namelist()
        for filename in sorted(file_list):
            if (filename.startswith("word/header") or filename.startswith("word/footer")) and filename.endswith(".xml"):
                xml_parts.append(z.read(filename).decode("utf-8", errors="ignore"))

    xml = "\n".join(xml_parts)

    # normalize: drop ignorable attrs and extra whitespace
    xml = IGNORABLE_PATTERNS.sub("", xml)
    xml = re.sub(r">\s+<", "><", xml)
    return xml


def read_comments_xml(docx_path: str) -> str:
    """
    Read and normalize comments-related XML parts from a .docx file.

    Includes:
    - word/comments.xml
    - word/commentsExtended.xml

    Returns:
        Combined normalized content of comments XML files
    """
    xml_parts = []

    with zipfile.ZipFile(docx_path, "r") as z:
        file_list = z.namelist()
        for filename in sorted(file_list):
            if filename in ("word/comments.xml", "word/commentsExtended.xml"):
                xml_parts.append(z.read(filename).decode("utf-8", errors="ignore"))

    xml = "\n".join(xml_parts)

    # normalize: drop ignorable attrs and extra whitespace
    xml = IGNORABLE_PATTERNS.sub("", xml)
    xml = re.sub(r">\s+<", "><", xml)
    return xml


def compare_doc_in_gold(doc_in_path: str, gold_path: str, xml_parts: str = "main") -> Tuple[bool, str]:
    """
    Compare doc_in with gold to detect if command changed the document.

    Args:
        doc_in_path: Path to doc_in.docx
        gold_path: Path to gold.docx
        xml_parts: Which XML parts to compare - "main", "supplementary", "footnotes", or "comments"

    Returns:
        (is_unchanged: bool, report: str)
    """
    if not os.path.exists(doc_in_path):
        return False, f"[ERROR] doc_in not found: {doc_in_path}"

    if not os.path.exists(gold_path):
        return False, f"[ERROR] gold not found: {gold_path}"

    try:
        if xml_parts == "main":
            doc_in_xml = read_main_xml(doc_in_path)
            gold_xml = read_main_xml(gold_path)
        elif xml_parts == "supplementary":
            doc_in_xml = read_supplementary_xml(doc_in_path)
            gold_xml = read_supplementary_xml(gold_path)
        elif xml_parts == "footnotes":
            doc_in_xml = read_footnotes_xml(doc_in_path)
            gold_xml = read_footnotes_xml(gold_path)
        elif xml_parts == "comments":
            doc_in_xml = read_comments_xml(doc_in_path)
            gold_xml = read_comments_xml(gold_path)
        else:
            return False, f"[ERROR] Invalid xml_parts value: {xml_parts}"

    except KeyError as e:
        return False, f"[ERROR] Missing part in DOCX: {e}"
    except Exception as e:
        return False, f"[ERROR] Failed to read DOCX: {e}"

    if doc_in_xml == gold_xml:
        return True, "[NO_XML_EFFECT_EXECUTION] doc_in equals gold (no document change)"
    else:
        return False, "[XML_EFFECT_EXECUTION] doc_in differs from gold (document modified)"


# ---------- Benchmark directory scanning ----------

def find_all_test_cases(benchmark_dir: Path) -> List[Dict]:
    """
    Scan benchmark directory and find all test cases with doc_in and gold files.

    Expected structure (content_formatting/content_insert):
        benchmark_dir/
            CommandName/
                variation_name/
                    doc_in.docx
                    gold.docx
                    spec.json

    Expected structure (object_effect):
        benchmark_dir/
            inputs/
                doc_in_single_shape.docx
                doc_in_two_shapes.docx
                doc_in_grouped_shape.docx
            CommandName/
                variation_name/
                    gold.docx
                    spec.json (with doc_in reference)

    Returns:
        List of test case dicts with paths
    """
    test_cases = []

    if not benchmark_dir.exists():
        print(f"[ERROR] Benchmark directory not found: {benchmark_dir}")
        return test_cases

    # Iterate through command directories
    for command_dir in benchmark_dir.iterdir():
        if not command_dir.is_dir():
            continue

        # Skip inputs directory
        if command_dir.name == 'inputs':
            continue

        command_name = command_dir.name

        # Iterate through variation directories
        for variation_dir in command_dir.iterdir():
            if not variation_dir.is_dir():
                continue

            variation_name = variation_dir.name
            doc_in_path = variation_dir / 'doc_in.docx'
            gold_path = variation_dir / 'gold.docx'
            spec_path = variation_dir / 'spec.json'

            # Check if gold exists (required)
            if not gold_path.exists():
                continue

            # Read spec.json to get configuration
            xml_parts = "main"  # default
            if not doc_in_path.exists() and spec_path.exists():
                try:
                    with open(spec_path, 'r', encoding='utf-8') as f:
                        spec = json.load(f)

                    # Check if spec has doc_in reference
                    doc_in_ref = spec.get('files', {}).get('doc_in')
                    if doc_in_ref:
                        # Resolve relative path from spec location
                        # Use Path.resolve() only for existence check;
                        # normalize with parent traversal but keep relative
                        resolved = (variation_dir / doc_in_ref).resolve()
                        # Convert back to relative path from CWD to avoid
                        # hardcoding absolute paths in test_set_summary.json
                        try:
                            doc_in_path = resolved.relative_to(Path.cwd())
                        except ValueError:
                            # If resolved path is outside CWD, keep absolute
                            doc_in_path = resolved

                    # Read xml_parts setting (infer from verification method if not explicit)
                    method = spec.get('verification', {}).get('primary', {}).get('method', '')
                    if method == 'header_footer_content_check':
                        xml_parts = 'supplementary'
                    elif method == 'comment_resolve_check':
                        xml_parts = 'comments'
                    else:
                        xml_parts = spec.get('xml_parts', 'main')
                except Exception as e:
                    print(f"[WARNING] Failed to read spec for {command_name}/{variation_name}: {e}")
                    continue
            elif spec_path.exists():
                # Read xml_parts even if doc_in exists
                try:
                    with open(spec_path, 'r', encoding='utf-8') as f:
                        spec = json.load(f)
                    method = spec.get('verification', {}).get('primary', {}).get('method', '')
                    if method == 'header_footer_content_check':
                        xml_parts = 'supplementary'
                    elif method == 'comment_resolve_check':
                        xml_parts = 'comments'
                    else:
                        xml_parts = spec.get('xml_parts', 'main')
                except Exception as e:
                    print(f"[WARNING] Failed to read xml_parts from spec for {command_name}/{variation_name}: {e}")

            # Check if required files exist
            if doc_in_path.exists() and gold_path.exists():
                test_case = {
                    'idMso': command_name,
                    'variation': variation_name,
                    'doc_in': str(doc_in_path),
                    'gold': str(gold_path),
                    'spec': str(spec_path) if spec_path.exists() else None,
                    'xml_parts': xml_parts,
                }
                test_cases.append(test_case)

    return test_cases


def summarize_all_test_cases(test_cases: List[Dict]) -> Dict:
    """
    Summarize all test cases and categorize them.

    Returns:
        Dict with summary results and statistics
    """
    results = {
        'total': len(test_cases),
        'no_xml_effect_execution': 0,
        'xml_effect_execution': 0,
        'error': 0,
        'no_xml_effect_execution_commands': [],
        'xml_effect_execution_commands': [],
        'error_commands': [],
        'details': []
    }

    for i, test_case in enumerate(test_cases, 1):
        idMso = test_case['idMso']
        variation = test_case['variation']

        print(f"[{i}/{len(test_cases)}] Summarizing {idMso}/{variation}...", end=' ')

        is_unchanged, report = compare_doc_in_gold(
            test_case['doc_in'],
            test_case['gold'],
            test_case['xml_parts']
        )

        # Classify result
        if '[ERROR]' in report:
            status = 'error'
            results['error'] += 1
            results['error_commands'].append({
                'idMso': idMso,
                'variation': variation,
                'report': report,
                'paths': test_case
            })
        elif is_unchanged:
            status = 'no_xml_effect_execution'
            results['no_xml_effect_execution'] += 1
            results['no_xml_effect_execution_commands'].append({
                'idMso': idMso,
                'variation': variation,
                'report': report,
                'paths': test_case
            })
        else:
            status = 'xml_effect_execution'
            results['xml_effect_execution'] += 1
            results['xml_effect_execution_commands'].append({
                'idMso': idMso,
                'variation': variation,
                'report': report,
                'paths': test_case
            })

        print(f"{status.upper()}")

        # Add to details
        results['details'].append({
            'idMso': idMso,
            'variation': variation,
            'status': status,
            'report': report,
            'paths': test_case
        })

    return results


# ---------- CLI ----------

def main():
    parser = argparse.ArgumentParser(
        description='Summarize generated benchmark test cases and collect statistics'
    )
    parser.add_argument(
        '--benchmark-dir', '-b',
        required=True,
        help='Path to benchmark directory (e.g., ./benchmark/content_formatting_x_core_actionable)'
    )
    parser.add_argument(
        '--output', '-o',
        default=None,
        help='Output directory for results (defaults to benchmark-dir)'
    )

    args = parser.parse_args()

    benchmark_dir = Path(args.benchmark_dir)
    output_dir = Path(args.output) if args.output else benchmark_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n=== Test Case Summary ===")
    print(f"Benchmark directory: {benchmark_dir}")
    print(f"Output directory: {output_dir}\n")

    # Step 1: Find all test cases
    print("Scanning benchmark directory...")
    test_cases = find_all_test_cases(benchmark_dir)
    print(f"Found {len(test_cases)} test cases\n")

    if not test_cases:
        print("[WARNING] No test cases found. Exiting.")
        return

    # Step 2: Summarize all test cases
    print("Summarizing test cases...\n")
    results = summarize_all_test_cases(test_cases)

    # Step 3: Calculate statistics
    timestamp = datetime.now().isoformat()

    # Count unique commands
    unique_commands = set(tc['idMso'] for tc in test_cases)
    total_commands = len(unique_commands)
    total_test_cases = len(test_cases)

    # Print summary
    print(f"\n{'='*80}")
    print(f"Summary Complete")
    print(f"{'='*80}")
    print(f"Total Commands: {total_commands}")
    print(f"Total Test Cases: {total_test_cases}")
    print(f"  XML Effect Execution: {results['xml_effect_execution']} (document modified)")
    print(f"  No XML Effect Execution: {results['no_xml_effect_execution']} (needs review)")
    print(f"  Errors: {results['error']}")
    print(f"{'='*80}")

    # Step 4: Build hierarchical test_cases structure
    # Structure: { "Bold": { "on": {...}, "off": {...} }, ... }
    test_cases_hierarchical = {}
    for detail in results['details']:
        idMso = detail['idMso']
        variation = detail['variation']

        if idMso not in test_cases_hierarchical:
            test_cases_hierarchical[idMso] = {}

        test_cases_hierarchical[idMso][variation] = {
            'status': detail['status'],
            'doc_in': detail['paths'].get('doc_in', ''),
            'gold': detail['paths'].get('gold', ''),
        }

    # Build final summary output
    summary_output = {
        'summarized_at': timestamp,
        'test_set': benchmark_dir.name,
        'benchmark_dir': str(benchmark_dir),
        'statistics': {
            'total_commands': total_commands,
            'total_test_cases': total_test_cases,
            'xml_effect_execution': results['xml_effect_execution'],
            'no_xml_effect_execution': results['no_xml_effect_execution'],
            'errors': results['error'],
        },
        'test_cases': test_cases_hierarchical
    }

    # Save summary as JSON
    summary_json_path = output_dir / 'test_set_summary.json'
    with open(summary_json_path, 'w', encoding='utf-8') as f:
        json.dump(summary_output, f, indent=2, ensure_ascii=False, sort_keys=True)

    print(f"\nSummary saved to: {summary_json_path}")

    # Print command lists for reference
    if results['no_xml_effect_execution'] > 0:
        print(f"\nNo XML Effect Execution commands ({results['no_xml_effect_execution']}):")
        for cmd in results['no_xml_effect_execution_commands']:
            print(f"  - {cmd['idMso']}/{cmd['variation']}")

    if results['error'] > 0:
        print(f"\nError commands ({results['error']}):")
        for cmd in results['error_commands']:
            print(f"  - {cmd['idMso']}/{cmd['variation']}")


if __name__ == '__main__':
    main()
