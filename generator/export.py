"""
Export verified benchmark test cases for sharing.

This script creates a clean, shareable benchmark structure from the generated
test cases, filtering to only include verified working cases (xml_effect_execution).

Usage:
    python -m generator.export -t content_formatting_x_core_actionable
    python -m generator.export -t object_effect_x_core_actionable --output ./my_benchmark

Input requirements:
    - benchmark/{test_set}/test_set_summary.json (from summarize_test_cases.py)
    - instructions/{test_set}.json

Output structure:
    wordeval/
    ├── {test_set}/
    │   ├── test_cases.json      # Hierarchical index with instructions
    │   ├── inputs/              # Shared input files
    │   └── {Command}/{variation}/
    │       ├── doc_in.docx
    │       └── gold.docx
    └── results/                 # Default directory for predictions
        └── {test_set}/
            └── {Command}/{variation}/
                └── pred.docx    # (user creates these)

Note: To generate evaluate.py, use: python -m evaluation.export
"""

import json
import shutil
import argparse
import re
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
from pathlib import Path
from datetime import datetime
from typing import Dict, Optional, Any

from .config import DOCINFO_AUTHOR_FIXTURE, FILL_VALUES


NS_CORE_CP = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
NS_DC = 'http://purl.org/dc/elements/1.1/'
NS_DCTERMS = 'http://purl.org/dc/terms/'
NS_XSI = 'http://www.w3.org/2001/XMLSchema-instance'
NS_APP = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'
NS_W15 = 'http://schemas.microsoft.com/office/word/2012/wordml'

ANONYMIZED_DOCUMENT_AUTHOR = 'Anonymous Author'
ANONYMIZED_REVIEW_AUTHOR = 'Anonymous Reviewer'
ANONYMIZED_TIMESTAMP = '2000-01-01T00:00:00Z'

TEST_AUTHOR_RE = re.compile(r'^(user\d+|TestUser)$', re.IGNORECASE)
XML_AUTHOR_ATTR_RE = re.compile(rb'([A-Za-z_][\w.-]*:author=")([^"]*)(")')
XML_INITIALS_ATTR_RE = re.compile(rb'([A-Za-z_][\w.-]*:initials=")[^"]*(")')
XML_DATE_ATTR_RE = re.compile(rb'([A-Za-z_][\w.-]*:date=")[^"]*(")')
OFFICE_PACKAGE_EXTENSIONS = {'.docx', '.docm', '.xlsx', '.xlsm', '.pptx', '.pptm'}

ET.register_namespace('cp', NS_CORE_CP)
ET.register_namespace('dc', NS_DC)
ET.register_namespace('dcterms', NS_DCTERMS)
ET.register_namespace('xsi', NS_XSI)
ET.register_namespace('', NS_APP)


def _is_test_author(value: Optional[str]) -> bool:
    return bool(value and TEST_AUTHOR_RE.match(value))


def _doc_property_fields(spec: Optional[Dict]) -> set:
    """Collect document-property field names from a case spec."""
    fields = set()

    def visit(value: Any):
        if isinstance(value, dict):
            field_mapping = value.get('field_mapping')
            if isinstance(field_mapping, dict):
                fields.update(str(key) for key in field_mapping.keys())

            field = value.get('field')
            if isinstance(field, str):
                fields.add(field)

            for item in value.values():
                visit(item)
        elif isinstance(value, list):
            for item in value:
                visit(item)

    visit(spec)
    return fields


def _extract_expected_field_values(spec: Optional[Dict]) -> Dict[str, str]:
    """Pull first-seen field_mapping values for identity-bearing fields."""
    result: Dict[str, str] = {}
    if not spec:
        return result

    def visit(value: Any) -> None:
        if isinstance(value, dict):
            fm = value.get('field_mapping')
            if isinstance(fm, dict):
                for key in ('Author', 'Manager', 'Company'):
                    val = fm.get(key)
                    if key not in result and isinstance(val, str) and val:
                        result[key] = val
            for item in value.values():
                visit(item)
        elif isinstance(value, list):
            for item in value:
                visit(item)

    visit(spec)
    return result


def _build_docx_identity_context(
    full_spec: Optional[Dict],
    idMso: str,
    *,
    artifact_role: str,
) -> Dict:
    """Describe identity fields that are semantic for one exported artifact.

    Final-state ``field_mapping`` values belong to the gold document, not to
    the input.  Treating one context as both roles can pre-populate a requested
    Author/Company operation in ``doc_in.docx``.  Inputs preserve identity only
    when their precondition explicitly declares ``input_field_mapping`` (or the
    legacy ``precondition.field_mapping`` form).
    """
    if artifact_role not in {'input', 'gold'}:
        raise ValueError(f'Unsupported artifact role: {artifact_role}')

    full_spec = full_spec or {}
    precondition = (full_spec or {}).get('precondition') or {}
    if artifact_role == 'gold':
        fields = _doc_property_fields(full_spec)
        expected = _extract_expected_field_values(full_spec)
    else:
        input_mapping = (
            full_spec.get('input_field_mapping')
            or precondition.get('field_mapping')
            or {}
        )
        input_spec = {'field_mapping': input_mapping}
        fields = _doc_property_fields(input_spec)
        expected = _extract_expected_field_values(input_spec)

    # DocInfoAuthor intentionally reads the current input property.  Normalize
    # it to the benchmark's fictional value instead of preserving an OS user.
    if idMso == 'DocInfoAuthor':
        fields.add('Author')
        expected.setdefault('Author', DOCINFO_AUTHOR_FIXTURE)
    elif artifact_role == 'input' and idMso in {
        'DocInfoDocTitle', 'DocInfoFileName',
    }:
        # These sibling cases retain the long-standing shared fictional input.
        # DocInfoAuthor now has its own isolated fixture, but that isolation
        # must not rewrite the unrelated shared input from Megan to Anonymous.
        fields.add('Author')
        expected.setdefault('Author', FILL_VALUES['Author'])

    return {
        'preserve_creator': 'Author' in fields,
        'preserve_manager': 'Manager' in fields,
        'preserve_company': 'Company' in fields,
        'expected_creator': expected.get('Author'),
        'expected_manager': expected.get('Manager'),
        'expected_company': expected.get('Company'),
    }


def _merge_docx_identity_context(existing: Optional[Dict], new: Dict) -> Dict:
    if not existing:
        return dict(new)
    return {
        'preserve_creator': existing.get('preserve_creator', False) or new.get('preserve_creator', False),
        'preserve_manager': existing.get('preserve_manager', False) or new.get('preserve_manager', False),
        'preserve_company': existing.get('preserve_company', False) or new.get('preserve_company', False),
        'expected_creator': existing.get('expected_creator') or new.get('expected_creator'),
        'expected_manager': existing.get('expected_manager') or new.get('expected_manager'),
        'expected_company': existing.get('expected_company') or new.get('expected_company'),
    }


def _package_text_references(entries: Dict[str, bytes], value: Optional[str]) -> bool:
    """
    Return whether a package value appears outside identity-only parts.

    This preserves real test content, such as document-property fields rendered
    into the body/header/footer, while still allowing author metadata to be
    scrubbed when it only appears in package identity records.
    """
    if not value:
        return False

    encoded = value.encode('utf-8')
    identity_only_parts = {
        'docProps/core.xml',
        'docProps/app.xml',
        'word/people.xml',
        'word/comments.xml',
        'word/commentsExtended.xml',
    }
    for name, data in entries.items():
        if name in identity_only_parts:
            continue
        if encoded in data:
            return True
    return False


def _set_or_create(root: ET.Element, namespace: str, tag: str, text: str) -> None:
    child = root.find(f'{{{namespace}}}{tag}')
    if child is None:
        child = ET.SubElement(root, f'{{{namespace}}}{tag}')
    child.text = text


def _sanitize_core_xml(data: bytes, entries: Dict[str, bytes], context: Dict) -> bytes:
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return data

    creator = root.find(f'{{{NS_DC}}}creator')
    creator_text = creator.text if creator is not None else ''
    preserve_creator = (
        context.get('preserve_creator', False)
        or _package_text_references(entries, creator_text)
    )

    if preserve_creator:
        # When the test expects a specific creator value, normalize to it so a
        # real-person name on the source file (e.g., the OS user who generated
        # the docx) does not survive into the released benchmark.
        expected_creator = context.get('expected_creator')
        if expected_creator and creator_text != expected_creator:
            _set_or_create(root, NS_DC, 'creator', expected_creator)
    else:
        _set_or_create(root, NS_DC, 'creator', ANONYMIZED_DOCUMENT_AUTHOR)

    _set_or_create(root, NS_CORE_CP, 'lastModifiedBy', ANONYMIZED_DOCUMENT_AUTHOR)

    for tag in ('created', 'modified'):
        child = root.find(f'{{{NS_DCTERMS}}}{tag}')
        if child is not None:
            child.text = ANONYMIZED_TIMESTAMP

    return ET.tostring(root, encoding='utf-8', xml_declaration=True)


def _sanitize_app_xml(data: bytes, entries: Dict[str, bytes], context: Dict) -> bytes:
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return data

    manager = root.find(f'{{{NS_APP}}}Manager')
    if manager is not None and manager.text:
        manager_text = manager.text
        preserve_manager = (
            context.get('preserve_manager', False)
            or _package_text_references(entries, manager_text)
        )
        if preserve_manager:
            expected_manager = context.get('expected_manager')
            if expected_manager and manager_text != expected_manager:
                manager.text = expected_manager
        else:
            manager.text = ANONYMIZED_DOCUMENT_AUTHOR

    company = root.find(f'{{{NS_APP}}}Company')
    if company is not None and company.text:
        company_text = company.text
        preserve_company = (
            context.get('preserve_company', False)
            or _package_text_references(entries, company_text)
        )
        if preserve_company:
            expected_company = context.get('expected_company')
            if expected_company and company_text != expected_company:
                company.text = expected_company
        else:
            company.text = ''

    return ET.tostring(root, encoding='utf-8', xml_declaration=True)


def _sanitize_author_attrs(data: bytes) -> bytes:
    def replace_author(match: re.Match) -> bytes:
        author = match.group(2).decode('utf-8', errors='replace')
        if _is_test_author(author):
            return match.group(0)
        return match.group(1) + ANONYMIZED_REVIEW_AUTHOR.encode('utf-8') + match.group(3)

    data = XML_AUTHOR_ATTR_RE.sub(replace_author, data)
    data = XML_INITIALS_ATTR_RE.sub(
        lambda match: match.group(1) + ANONYMIZED_REVIEW_AUTHOR[:1].encode('utf-8') + match.group(2),
        data,
    )
    data = XML_DATE_ATTR_RE.sub(
        lambda match: match.group(1) + ANONYMIZED_TIMESTAMP.encode('utf-8') + match.group(2),
        data,
    )
    return data


def _empty_people_xml() -> bytes:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w15:people xmlns:w15="{NS_W15}"/>'
    ).encode('utf-8')


def _write_zip_entries(target: Any, entries: Dict[str, bytes], infos: Dict[str, zipfile.ZipInfo]) -> None:
    with zipfile.ZipFile(target, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries.items():
            info = zipfile.ZipInfo(name)
            source_info = infos.get(name)
            if source_info is not None:
                info.external_attr = source_info.external_attr
                info.compress_type = zipfile.ZIP_DEFLATED
            info.date_time = (2000, 1, 1, 0, 0, 0)
            zout.writestr(info, data)


def _sanitize_office_package_bytes(data: bytes, context: Dict, depth: int = 0) -> bytes:
    """Sanitize an embedded Office zip package, such as a chart workbook."""
    try:
        with zipfile.ZipFile(BytesIO(data), 'r') as zin:
            entries = {item.filename: zin.read(item.filename) for item in zin.infolist()}
            infos = {item.filename: item for item in zin.infolist()}
    except zipfile.BadZipFile:
        return data

    for name, item_data in list(entries.items()):
        if name == 'docProps/core.xml':
            entries[name] = _sanitize_core_xml(item_data, entries, context)
        elif name == 'docProps/app.xml':
            entries[name] = _sanitize_app_xml(item_data, entries, context)
        elif name == 'word/people.xml':
            entries[name] = _empty_people_xml()
        elif name.endswith('.xml'):
            entries[name] = _sanitize_author_attrs(item_data)
        elif depth < 2 and Path(name).suffix.lower() in OFFICE_PACKAGE_EXTENSIONS:
            entries[name] = _sanitize_office_package_bytes(item_data, context, depth + 1)

    output = BytesIO()
    _write_zip_entries(output, entries, infos)
    return output.getvalue()


def sanitize_docx_identity(src: Path, dest: Path, context: Optional[Dict] = None) -> None:
    """
    Copy a docx while removing author/account identity from package metadata.

    The sanitizer keeps semantic document-property values when the case spec or
    rendered package content indicates those values are part of the test.
    """
    context = context or {}
    dest.parent.mkdir(parents=True, exist_ok=True)

    if src.suffix.lower() != '.docx':
        shutil.copy2(src, dest)
        return

    with zipfile.ZipFile(src, 'r') as zin:
        entries = {item.filename: zin.read(item.filename) for item in zin.infolist()}
        infos = {item.filename: item for item in zin.infolist()}

    for name, data in list(entries.items()):
        if name == 'docProps/core.xml':
            entries[name] = _sanitize_core_xml(data, entries, context)
        elif name == 'docProps/app.xml':
            entries[name] = _sanitize_app_xml(data, entries, context)
        elif name == 'word/people.xml':
            entries[name] = _empty_people_xml()
        elif name.endswith('.xml') and name.startswith('word/'):
            entries[name] = _sanitize_author_attrs(data)
        elif Path(name).suffix.lower() in OFFICE_PACKAGE_EXTENSIONS:
            entries[name] = _sanitize_office_package_bytes(data, context)

    _write_zip_entries(dest, entries, infos)


def load_test_set_summary(benchmark_dir: Path) -> Optional[Dict]:
    """Load test set summary from the benchmark directory."""
    summary_path = benchmark_dir / 'test_set_summary.json'
    if not summary_path.exists():
        print(f"[ERROR] Test set summary not found: {summary_path}")
        print("  Run summarize_test_cases.py first to generate summary.")
        return None

    with open(summary_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def load_instructions(instructions_dir: Path, test_set: str) -> Optional[Dict]:
    """Load instructions from the instructions directory."""
    instructions_path = instructions_dir / f'{test_set}.json'
    if not instructions_path.exists():
        print(f"[ERROR] Instructions not found: {instructions_path}")
        return None

    with open(instructions_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def load_instructions_from_variants(instructions_dir: Path, test_set: str) -> Optional[Dict]:
    """
    Load instructions from variant files for parameterized test sets.

    This is used for test sets where instructions are defined in variant files
    (single source of truth) rather than duplicated in main instructions file.
    """
    # Find all variant files for this test set
    # Match both {test_set}_variants.json and {test_set}_*_variants.json
    variant_files = list(instructions_dir.glob(f'{test_set}_*_variants.json'))
    direct_file = instructions_dir / f'{test_set}_variants.json'
    if direct_file.exists() and direct_file not in variant_files:
        variant_files.append(direct_file)

    if not variant_files:
        print(f"[ERROR] No variant files found for {test_set}")
        return None

    # Try to read description from main instructions file
    main_instructions_path = instructions_dir / f'{test_set}.json'
    description = f'Instructions loaded from variant files for {test_set}'
    if main_instructions_path.exists():
        with open(main_instructions_path, 'r', encoding='utf-8') as f:
            main_data = json.load(f)
        if main_data.get('description', ''):
            description = main_data['description']

    # Build instructions dict from variant files
    instructions = {
        'test_set': test_set,
        'description': description,
        'commands': {}
    }

    for variant_file in variant_files:
        with open(variant_file, 'r', encoding='utf-8') as f:
            variant_data = json.load(f)

        for cmd_name, cmd_data in variant_data.get('commands', {}).items():
            if cmd_name not in instructions['commands']:
                instructions['commands'][cmd_name] = {}

            for variant in cmd_data.get('variants', []):
                # Only include variants marked for generation
                if not variant.get('generate', True):
                    continue

                label = variant.get('label', '')
                instruction = variant.get('instruction', '')

                if instruction:
                    # Store under both 'apply_{label}' and '{label}' to match
                    # whichever naming convention the test set uses for variations
                    instructions['commands'][cmd_name][f'apply_{label}'] = instruction
                    instructions['commands'][cmd_name][label] = instruction

    return instructions


def load_instructions_with_variant_fallback(
    instructions_dir: Path,
    test_set: str,
) -> Optional[Dict]:
    """Load main instructions and fill missing variations from variant files."""
    instructions = load_instructions(instructions_dir, test_set)
    variant_instructions = load_instructions_from_variants(instructions_dir, test_set)

    if not instructions:
        return variant_instructions
    if not variant_instructions:
        return instructions

    commands = instructions.setdefault('commands', {})
    for cmd_name, variant_cmd in variant_instructions.get('commands', {}).items():
        cmd_instructions = commands.setdefault(cmd_name, {})
        for variation, instruction in variant_cmd.items():
            cmd_instructions.setdefault(variation, instruction)

    return instructions


# Test sets that read instructions from variant files (single source of truth)
VARIANT_BASED_TEST_SETS = {
    'content_formatting_x_core_parameterized',
    'content_operations_x_core_parameterized',
    'object_effect_x_core_parameterized',
}

# Test sets with curated main instructions plus variant-file fallback.
VARIANT_FALLBACK_TEST_SETS = {
    'object_effect_x_contextual_parameterized',
}


def get_instruction(instructions: Dict, idMso: str, variation: str) -> str:
    """Get instruction for a specific command and variation."""
    commands = instructions.get('commands', {})
    cmd_instructions = commands.get(idMso, {})
    return cmd_instructions.get(variation, f"Apply {idMso} ({variation}).")


def export_benchmark(
    test_set: str,
    benchmark_dir: Path,
    instructions_dir: Path,
    output_dir: Path,
    include_no_effect: bool = False,
    case_filters: Optional[set[tuple[str, str]]] = None,
) -> bool:
    """
    Export verified benchmark test cases.

    Args:
        test_set: Name of the test set (e.g., 'content_formatting_x_core_actionable')
        benchmark_dir: Path to benchmark/{test_set}
        instructions_dir: Path to instructions/
        output_dir: Path to output directory (e.g., wordeval)
        include_no_effect: If True, also include no_xml_effect_execution cases

    Returns:
        True if successful, False otherwise
    """
    # Load test set summary
    summary = load_test_set_summary(benchmark_dir)
    if not summary:
        return False

    # Load instructions (from variant files for parameterized test sets)
    if test_set in VARIANT_BASED_TEST_SETS:
        instructions = load_instructions_from_variants(instructions_dir, test_set)
        if instructions:
            print(f"  [INFO] Loaded instructions from variant files (single source of truth)")
    elif test_set in VARIANT_FALLBACK_TEST_SETS:
        instructions = load_instructions_with_variant_fallback(instructions_dir, test_set)
        if instructions:
            print(f"  [INFO] Loaded main instructions with variant-file fallback")
    else:
        instructions = load_instructions(instructions_dir, test_set)

    if not instructions:
        return False

    # Get test cases from summary
    test_cases_raw = summary.get('test_cases', {})

    # Filter to valid statuses
    valid_statuses = ['xml_effect_execution']
    if include_no_effect:
        valid_statuses.append('no_xml_effect_execution')

    # Build hierarchical test_cases structure for output
    test_cases_output = {}
    exported_count = 0
    skipped_count = 0

    # Track which input files we need to copy
    input_files_to_copy = set()
    input_file_contexts = {}

    # Test set output directory
    test_set_output_dir = output_dir / test_set

    for idMso, variations in test_cases_raw.items():
        for variation, data in variations.items():
            if case_filters and (idMso, variation) not in case_filters:
                continue
            status = data.get('status', '')

            if status not in valid_statuses:
                skipped_count += 1
                continue

            # Read spec.json as the authoritative source for file paths
            # (test_set_summary.json paths can go stale if spec is updated
            # without re-running summarize)
            spec_path = benchmark_dir / idMso / variation / 'spec.json'
            variation_dir = benchmark_dir / idMso / variation
            full_spec = None
            verification_spec = None

            if spec_path.exists():
                try:
                    with open(spec_path, 'r', encoding='utf-8') as f:
                        full_spec = json.load(f)
                    verification_spec = full_spec.get('verification')
                except Exception as e:
                    print(f"[WARNING] Failed to read {spec_path}: {e}")

            # Resolve file paths from spec.json (authoritative)
            preview_path = None
            if full_spec and 'files' in full_spec:
                spec_files = full_spec['files']
                doc_in_path = (variation_dir / spec_files['doc_in']).resolve()
                gold_path = (variation_dir / spec_files['gold']).resolve()
                preview_rel = spec_files.get('preview')
                if preview_rel:
                    candidate = (variation_dir / preview_rel).resolve()
                    if candidate.exists():
                        preview_path = candidate
            else:
                # Fallback: try paths from summary (legacy test sets without spec)
                doc_in_path = variation_dir / 'doc_in.docx'
                gold_path = variation_dir / 'gold.docx'

            if not doc_in_path.exists() or not gold_path.exists():
                print(f"[WARNING] Missing files for {idMso}/{variation}, skipping")
                if not doc_in_path.exists():
                    print(f"  doc_in not found: {doc_in_path}")
                if not gold_path.exists():
                    print(f"  gold not found: {gold_path}")
                skipped_count += 1
                continue

            # Get instruction
            instruction = get_instruction(instructions, idMso, variation)
            input_identity_context = _build_docx_identity_context(
                full_spec, idMso, artifact_role='input')
            gold_identity_context = _build_docx_identity_context(
                full_spec, idMso, artifact_role='gold')

            # Determine relative doc_in path for output
            doc_in_relative = None
            if 'inputs' in str(doc_in_path):
                input_files_to_copy.add(doc_in_path)
                input_file_contexts[doc_in_path] = _merge_docx_identity_context(
                    input_file_contexts.get(doc_in_path),
                    input_identity_context,
                )
                doc_in_relative = f"inputs/{doc_in_path.name}"
            else:
                doc_in_relative = f"{idMso}/{variation}/doc_in.docx"

            # Collect auxiliary files (e.g., img.png for picture prerequisite)
            if full_spec and 'files' in full_spec:
                for aux_rel in full_spec['files'].get('auxiliary', []):
                    aux_path = (variation_dir / aux_rel).resolve()
                    if aux_path.exists() and 'inputs' in str(aux_path):
                        input_files_to_copy.add(aux_path)
                        input_file_contexts[aux_path] = _merge_docx_identity_context(
                            input_file_contexts.get(aux_path),
                            input_identity_context,
                        )

            # Build test case entry
            if idMso not in test_cases_output:
                test_cases_output[idMso] = {}

            entry = {
                'instruction': instruction,
                'doc_in': doc_in_relative,
                'gold': f"{idMso}/{variation}/gold.docx",
                '_source': {
                    'doc_in_abs': str(doc_in_path),
                    'gold_abs': str(gold_path),
                    'input_identity_context': input_identity_context,
                    'gold_identity_context': gold_identity_context,
                }
            }
            if preview_path is not None:
                preview_name = preview_path.name
                entry['preview'] = f"{idMso}/{variation}/{preview_name}"
                entry['_source']['preview_abs'] = str(preview_path)
                # Add multimodal instruction with preview reference
                # (instruction stays as original text-only version)
                instr = entry['instruction']
                if '")' in instr:
                    entry['instruction_mm'] = instr.replace(
                        '")', '", as shown in @preview.png)', 1)
                else:
                    entry['instruction_mm'] = instr.rstrip('.') + ', as shown in @preview.png.'

            # Export the verification block, behavior, and precondition for verification
            if verification_spec:
                entry['spec'] = {'verification': verification_spec}
                # Include behavior so invariant checks can be behavior-aware
                if 'behavior' in full_spec:
                    entry['spec']['behavior'] = full_spec['behavior']
                # Include precondition so verifiers know target selection (e.g., "all", "paragraph_3")
                if 'precondition' in full_spec:
                    entry['spec']['precondition'] = full_spec['precondition']
                # Include composition metadata for composition test cases
                if 'composition' in full_spec:
                    composition = full_spec.get('composition') or {}
                    if 'commands' in composition:
                        raise ValueError(
                            f"{spec_path} contains legacy composition.commands; "
                            "regenerate this composition test set before export."
                        )
                    if not composition.get('atomic_intents'):
                        raise ValueError(
                            f"{spec_path} is a composition spec without "
                            "composition.atomic_intents; regenerate or fix the "
                            "composition generator before export."
                        )
                    entry['spec']['composition'] = composition

            test_cases_output[idMso][variation] = entry

            exported_count += 1

    if exported_count == 0:
        print(f"[ERROR] No valid test cases to export")
        return False

    # Create output directory structure
    test_set_output_dir.mkdir(parents=True, exist_ok=True)

    # Copy input files
    inputs_output_dir = test_set_output_dir / 'inputs'
    inputs_output_dir.mkdir(exist_ok=True)

    for input_path in input_files_to_copy:
        dest = inputs_output_dir / input_path.name
        sanitize_docx_identity(input_path, dest, input_file_contexts.get(input_path))


    # Copy test case files and remove _source from output
    for idMso, variations in test_cases_output.items():
        for variation, data in variations.items():
            # Create output directory
            case_output_dir = test_set_output_dir / idMso / variation
            case_output_dir.mkdir(parents=True, exist_ok=True)

            # Copy gold.docx
            gold_src = Path(data['_source']['gold_abs'])
            gold_dest = case_output_dir / 'gold.docx'
            sanitize_docx_identity(
                gold_src,
                gold_dest,
                data['_source'].get('gold_identity_context'),
            )

            # Copy preview image if present
            if 'preview_abs' in data['_source']:
                preview_src = Path(data['_source']['preview_abs'])
                if preview_src.exists():
                    preview_dest = case_output_dir / preview_src.name
                    shutil.copy2(preview_src, preview_dest)

            # Copy doc_in.docx if it's local (not in inputs/)
            if not data['doc_in'].startswith('inputs/'):
                doc_in_src = Path(data['_source']['doc_in_abs'])
                doc_in_dest = case_output_dir / 'doc_in.docx'
                sanitize_docx_identity(
                    doc_in_src,
                    doc_in_dest,
                    data['_source'].get('input_identity_context'),
                )

            # Remove _source from output data
            del data['_source']

    # Build final test_cases.json
    output_json = {
        'test_set': test_set,
        'version': '1.0',
        'exported_at': datetime.now().isoformat(),
        'description': instructions.get('description', ''),
        'statistics': {
            'total_exported': exported_count,
            'skipped': skipped_count,
        },
        'commands': test_cases_output,
    }

    # Save test_cases.json
    test_cases_path = test_set_output_dir / 'test_cases.json'
    if case_filters and test_cases_path.exists():
        with open(test_cases_path, 'r', encoding='utf-8') as f:
            existing_output = json.load(f)
        existing_commands = existing_output.setdefault('commands', {})
        for idMso, variations in test_cases_output.items():
            existing_commands.setdefault(idMso, {}).update(variations)
        existing_output['exported_at'] = output_json['exported_at']
        output_json = existing_output
    with open(test_cases_path, 'w', encoding='utf-8') as f:
        json.dump(output_json, f, indent=2, ensure_ascii=False, sort_keys=True)

    print(f"\n[SUCCESS] Exported {exported_count} test cases for {test_set}")
    print(f"  Skipped: {skipped_count} (no effect or errors)")

    return True


def main():
    parser = argparse.ArgumentParser(
        description='Export verified benchmark test cases for sharing'
    )
    parser.add_argument(
        '--test-set', '-t',
        action='append',
        dest='test_sets',
        required=True,
        help='Test set name (can be specified multiple times)'
    )
    parser.add_argument(
        '--benchmark', '-b',
        default=None,
        help='Path to benchmark directory (defaults to ./benchmark)'
    )
    parser.add_argument(
        '--instructions', '-i',
        default=None,
        help='Path to instructions directory (defaults to ./instructions)'
    )
    parser.add_argument(
        '--output', '-o',
        default=None,
        help='Output directory (defaults to ./wordeval)'
    )
    parser.add_argument(
        '--include-no-effect',
        action='store_true',
        help='Also include no_xml_effect_execution cases (for review)'
    )
    parser.add_argument(
        '--case', action='append', default=[], metavar='COMMAND/VARIATION',
        help='Export only one exact case and merge it into an existing test set',
    )

    args = parser.parse_args()

    # Determine paths
    base_dir = Path(__file__).parent.parent
    benchmark_base = Path(args.benchmark) if args.benchmark else base_dir / 'benchmark'
    instructions_dir = Path(args.instructions) if args.instructions else base_dir / 'instructions'
    output_dir = Path(args.output) if args.output else base_dir / 'wordeval'
    case_filters = set()
    for value in args.case:
        if '/' not in value:
            parser.error(f"--case must be COMMAND/VARIATION, got: {value}")
        command, variation = value.split('/', 1)
        if not command or not variation:
            parser.error(f"--case must be COMMAND/VARIATION, got: {value}")
        case_filters.add((command, variation))

    print(f"=== Export WordEval Test Cases ===")
    print(f"Output: {output_dir}")
    print()

    # Export each test set
    exported_sets = []
    for test_set in args.test_sets:
        test_set_benchmark_dir = benchmark_base / test_set

        print(f"Processing: {test_set}")
        print(f"  Source: {test_set_benchmark_dir}")

        if not test_set_benchmark_dir.exists():
            print(f"  [ERROR] Benchmark directory not found: {test_set_benchmark_dir}")
            continue

        success = export_benchmark(
            test_set=test_set,
            benchmark_dir=test_set_benchmark_dir,
            instructions_dir=instructions_dir,
            output_dir=output_dir,
            include_no_effect=args.include_no_effect,
            case_filters=case_filters or None,
        )

        if success:
            exported_sets.append(test_set)
        print()

    if not exported_sets:
        print("[ERROR] No test sets exported successfully")
        return 1

    print(f"\n=== Export Complete ===")
    print(f"Output directory: {output_dir}")
    print(f"Test sets: {', '.join(exported_sets)}")
    print(f"\nNext steps:")
    print(f"  1. Generate evaluate.py: python -m evaluation.export")
    print(f"  2. Evaluate results: python -m evaluation.cli -t {exported_sets[0]} -r results/")

    return 0


if __name__ == '__main__':
    exit(main())
