"""
Test case generator for content_operations × contextual_actionable.

This test set covers commands that live on a contextual tab (currently the
Header/Footer Tools tab) whose action is to insert document-info content
into the document. They were previously misclassified as object_effect
because the contextual tab pattern dominated, but their semantic effect is
content insertion.

Currently scoped to 3 commands, all single-variant ('apply'). Note that the
three commands produce *different OOXML artifacts*, mirroring real Word UI:

- DocInfoAuthor    — inserts a content control (SDT) into the header,
                     bound to coreProperties/dc:creator via dataBinding,
                     pre-filled with the current Author value.
- DocInfoDocTitle  — inserts a content control bound to dc:title, pre-filled
                     with the current Title value (same SDT mechanism).
- DocInfoFileName  — inserts a FILENAME field (wdFieldFileName, type=29).
                     There is no docproperty binding for "file name", so the
                     SDT path used by Author/Title does not apply here.

DocInfoFilePath is intentionally excluded for now. It inserts a FILENAME field
with the \\p full-path switch, whose cached rendered result can expose local
paths and vary across environments. It needs a field-code-only verifier and
artifact sanitization before release.

Each test case opens `inputs/doc_in_info.docx` (a document with Author/Title
properties pre-filled), enters header-edit mode, runs the command, then exits
and saves as gold.docx.
"""

import argparse
import json
import shutil
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

from ..config import (
    BENCHMARK_OUTPUT_DIR,
    DOCINFO_AUTHOR_FIXTURE,
    FILL_VALUES,
)
from ..execution.setup import ensure_word
from ..execution.execution import DirectExecute
from generator.test_sets.spec_utils import write_spec_if_changed


TEST_SET_NAME = 'content_operations_x_contextual_actionable'

# Commands handled here. Each maps to a single 'apply' variant.
DOC_INFO_COMMANDS = ('DocInfoAuthor', 'DocInfoDocTitle', 'DocInfoFileName')


def _builtin_document_property(doc, name: str):
    """Return the raw DocumentProperty dispatch across pywin32 wrappers."""
    return doc.BuiltInDocumentProperties._oleobj_.Invoke(
        0, 0, 2, 1, name  # DISPID_VALUE / DISPATCH_PROPERTYGET
    )


def _get_builtin_document_property(doc, name: str):
    return _builtin_document_property(doc, name).Invoke(0, 0, 2, 1)


def _set_builtin_document_property(doc, name: str, value) -> None:
    _builtin_document_property(doc, name).Invoke(0, 0, 4, 0, value)


def _setup_input_doc(doc, *, author_value: str = FILL_VALUES['Author']) -> None:
    """Set Author/Title properties so DocInfo fields render visible content."""
    try:
        _set_builtin_document_property(doc, "Author", author_value)
    except Exception:
        pass
    try:
        _set_builtin_document_property(doc, "Title", FILL_VALUES['Title'])
    except Exception:
        pass


def _enter_header_edit(wd) -> None:
    """Move the active view into the page-1 header so field insertion goes there."""
    wd.ActiveWindow.ActivePane.View.SeekView = 9  # wdSeekCurrentPageHeader
    time.sleep(0.3)


def _exit_header_edit(wd) -> None:
    wd.ActiveWindow.ActivePane.View.SeekView = 0  # wdSeekMainDocument
    time.sleep(0.2)


def _execute(wd, idMso: str) -> bool:
    """Run the DocInfo command. Author / DocTitle work via ExecuteMso. FileName
    needs a direct Fields.Add because ExecuteMso targets the wrong field type
    when invoked via COM."""
    selection = wd.Selection
    if idMso == 'DocInfoFileName':
        selection.Fields.Add(Range=selection.Range, Type=29)  # wdFieldFileName
        return True
    try:
        wd.CommandBars.ExecuteMso(idMso)
        time.sleep(0.2)
        return True
    except Exception as e:
        print(f"  ExecuteMso({idMso}) failed: {e}")
        return False


def _build_spec(idMso: str, input_filename: str = 'doc_in_info.docx') -> Dict:
    spec = {
        'idMso': idMso,
        'behavior': 'content_operations',
        'scope': 'contextual_actionable',
        'variation': 'apply',
        'generated_at': datetime.now().isoformat(),
        'precondition': {
            'type': 'info',
            'description': 'Document with Author/Title properties pre-filled, viewed in header edit mode',
        },
        'verification': {
            'primary': {
                'method': 'docinfo_verification',
                'description': 'Verify the semantic DocInfo field/content control in the header',
                'part_type': 'header',
            },
            'invariants': [
                {
                    'description': 'Body text content preserved',
                    'method': 'text_preserved',
                },
            ],
        },
        'files': {
            'doc_in': f'../../inputs/{input_filename}',
            'gold': 'gold.docx',
        },
    }
    if idMso == 'DocInfoAuthor':
        spec['precondition']['description'] = (
            "Document whose Author core property is the synthetic fixture "
            f"'{DOCINFO_AUTHOR_FIXTURE}', viewed in header edit mode"
        )
        spec['input_field_mapping'] = {
            'Author': DOCINFO_AUTHOR_FIXTURE,
        }
    return spec


def _validate_docinfo_author_snapshot(snapshot: Dict) -> None:
    """Validate the saved DocInfoAuthor source and mapped-control contract."""
    errors = []
    if snapshot.get('creator') != DOCINFO_AUTHOR_FIXTURE:
        errors.append(
            'Author core property mismatch: '
            f"{snapshot.get('creator')!r} != {DOCINFO_AUTHOR_FIXTURE!r}"
        )

    controls = snapshot.get('author_controls') or []
    if len(controls) != 1:
        errors.append(f'expected exactly one Author content control, found {len(controls)}')
    elif controls:
        control = controls[0]
        if not control.get('is_mapped'):
            errors.append('Author content control is not mapped')
        if not control.get('source_exists'):
            errors.append('Author content control mapping source is missing')
        if 'creator' not in str(control.get('xpath') or '').lower():
            errors.append('Author content control maps to the wrong core property')
        if control.get('text') != DOCINFO_AUTHOR_FIXTURE:
            errors.append(
                'Author content control text mismatch: '
                f"{control.get('text')!r} != {DOCINFO_AUTHOR_FIXTURE!r}"
            )

    if errors:
        raise ValueError('; '.join(errors))


def _author_snapshot(doc) -> Dict:
    """Read the Author property and header controls from an open Word document."""
    try:
        creator = str(_get_builtin_document_property(doc, 'Author') or '')
    except Exception:
        creator = ''

    controls = []
    seen_ids = set()
    for section_index in range(1, doc.Sections.Count + 1):
        section = doc.Sections(section_index)
        for header_index in range(1, section.Headers.Count + 1):
            header = section.Headers(header_index)
            try:
                if not header.Exists:
                    continue
                header_controls = header.Range.ContentControls
            except Exception:
                continue
            for control_index in range(1, header_controls.Count + 1):
                control = header_controls(control_index)
                try:
                    control_id = str(control.ID)
                except Exception:
                    control_id = f'{section_index}:{header_index}:{control_index}'
                if control_id in seen_ids:
                    continue
                seen_ids.add(control_id)
                try:
                    title = str(control.Title or '')
                except Exception:
                    title = ''
                try:
                    tag = str(control.Tag or '')
                except Exception:
                    tag = ''
                if title.lower() != 'author' and tag.lower() != 'author':
                    continue
                try:
                    mapping = control.XMLMapping
                    is_mapped = bool(mapping.IsMapped)
                    xpath = str(mapping.XPath or '')
                    source_exists = bool(is_mapped and mapping.CustomXMLPart is not None)
                except Exception:
                    is_mapped = False
                    xpath = ''
                    source_exists = False
                try:
                    text = str(control.Range.Text or '').strip('\r\n\x07 ')
                except Exception:
                    text = ''
                controls.append({
                    'title': title,
                    'tag': tag,
                    'text': text,
                    'is_mapped': is_mapped,
                    'source_exists': source_exists,
                    'xpath': xpath,
                })

    return {
        'creator': creator,
        'author_controls': controls,
    }


def _assert_docinfo_author_postconditions(wd, input_path: Path, gold_path: Path) -> None:
    """Normal-reopen the Author input/reference and enforce binding/source state."""
    input_doc = None
    gold_doc = None
    try:
        input_doc = wd.Documents.Open(
            str(input_path.absolute()), ReadOnly=True, AddToRecentFiles=False
        )
        input_snapshot = _author_snapshot(input_doc)
        if input_snapshot.get('creator') != DOCINFO_AUTHOR_FIXTURE:
            raise ValueError(
                'DocInfoAuthor input did not preserve the synthetic Author source'
            )
        input_doc.Close(False)
        input_doc = None

        gold_doc = wd.Documents.Open(
            str(gold_path.absolute()), ReadOnly=True, AddToRecentFiles=False
        )
        _validate_docinfo_author_snapshot(_author_snapshot(gold_doc))
    finally:
        if input_doc is not None:
            try:
                input_doc.Close(False)
            except Exception:
                pass
        if gold_doc is not None:
            try:
                gold_doc.Close(False)
            except Exception:
                pass


class ContentOperationsContextualActionableGenerator:
    def __init__(self, output_dir: Path) -> None:
        self.output_dir = output_dir
        self.inputs_dir = output_dir / 'inputs'
        self.execution = DirectExecute()

    def ensure_input(self, idMso: str) -> Path:
        """Generate the command's input, isolating DocInfoAuthor identity."""
        self.inputs_dir.mkdir(parents=True, exist_ok=True)
        is_author = idMso == 'DocInfoAuthor'
        input_filename = 'doc_in_author.docx' if is_author else 'doc_in_info.docx'
        input_path = self.inputs_dir / input_filename
        if input_path.exists() and not is_author:
            return input_path
        if input_path.exists():
            input_path.unlink()

        wd = ensure_word()
        doc = None
        try:
            doc = wd.Documents.Add()
            _setup_input_doc(
                doc,
                author_value=(
                    DOCINFO_AUTHOR_FIXTURE if is_author else FILL_VALUES['Author']
                ),
            )
            time.sleep(0.3)
            doc.SaveAs(str(input_path.absolute()), FileFormat=16)
            if is_author:
                # SaveAs seeds a new file's creator from the active Office
                # profile.  Set the synthetic source again on the saved
                # document and persist it so no real local identity leaks.
                _set_builtin_document_property(
                    doc, 'Author', DOCINFO_AUTHOR_FIXTURE
                )
                doc.Save()
            time.sleep(0.2)
        finally:
            if doc is not None:
                try:
                    doc.Close(False)
                except Exception:
                    pass
        return input_path

    def generate_one(self, idMso: str, input_path: Path) -> Dict:
        case_dir = self.output_dir / idMso / 'apply'
        case_dir.mkdir(parents=True, exist_ok=True)
        gold_path = case_dir / 'gold.docx'
        spec_path = case_dir / 'spec.json'

        result = {
            'idMso': idMso,
            'variation': 'apply',
            'success': False,
            'message': '',
        }

        wd = ensure_word()
        doc = None
        try:
            doc = wd.Documents.Open(str(input_path.absolute()))
            time.sleep(0.3)
            _enter_header_edit(wd)
            ok = _execute(wd, idMso)
            _exit_header_edit(wd)
            if not ok:
                result['message'] = f'execute failed: {idMso}'
                return result
            doc.SaveAs(str(gold_path.absolute()), FileFormat=16)
            time.sleep(0.2)

            doc.Close(False)
            doc = None
            if idMso == 'DocInfoAuthor':
                _assert_docinfo_author_postconditions(wd, input_path, gold_path)

            write_spec_if_changed(spec_path, _build_spec(idMso, input_path.name))
            result['success'] = True
            result['message'] = 'OK'
        except Exception as e:
            result['message'] = f'error: {e}'
            import traceback
            traceback.print_exc()
        finally:
            if doc is not None:
                try:
                    doc.Close(False)
                except Exception:
                    pass
        return result

    def run(self, only_command: Optional[str] = None) -> List[Dict]:
        cmds = [c for c in DOC_INFO_COMMANDS if (only_command is None or c == only_command)]
        results = []
        for cmd in cmds:
            input_path = self.ensure_input(cmd)
            print(f'Generating {cmd}/apply...')
            r = self.generate_one(cmd, input_path)
            print(f'  {r["message"]}')
            results.append(r)
        return results


def main() -> None:
    parser = argparse.ArgumentParser(description='Generate content_operations × contextual_actionable test cases')
    parser.add_argument('--command', help='Generate only this command', default=None)
    args, _ = parser.parse_known_args()

    output_dir = BENCHMARK_OUTPUT_DIR / TEST_SET_NAME
    output_dir.mkdir(parents=True, exist_ok=True)
    gen = ContentOperationsContextualActionableGenerator(output_dir)
    results = gen.run(only_command=args.command)

    summary_path = output_dir / 'generation_summary.json'
    summary = {
        'test_set': TEST_SET_NAME,
        'generated_at': datetime.now().isoformat(),
        'results': results,
        'total': len(results),
        'success': sum(1 for r in results if r['success']),
    }
    with open(summary_path, 'w', encoding='utf-8') as f:
        json.dump(summary, f, indent=2, ensure_ascii=False)
    print(f"\nSaved generation summary: {summary_path}")


if __name__ == '__main__':
    main()
