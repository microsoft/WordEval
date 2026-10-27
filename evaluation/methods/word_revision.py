"""Optional Word CompareDocuments baseline using shared document access.

Run through ``python -m experiments.word_revision`` for package validation,
owned Word processes, bounded execution and source-identified results.
"""

import os
from pathlib import Path
from collections import Counter
import traceback
import time
from ..document_io.word import WordSession, open_normal_document
from ..document_io.errors import NormalOpenError
from ..document_io.open_gate import normal_open_gate

# Revision type mapping (from Word VBA constants)
REVISION_TYPES = {
    1: "Insertion",
    2: "Deletion",
    3: "Property changed",
    4: "Paragraph number changed",
    5: "Field display changed",
    6: "Revision marked as reconciled conflict",
    7: "Revision marked as a conflict",
    8: "Style changed",
    9: "Replaced",
    10: "Paragraph property changed",
    11: "Table property changed",
    12: "Section property changed",
    13: "Style definition changed",
    14: "Content moved from",
    15: "Content moved to",
    16: "Table cell inserted",
    17: "Table cell deleted",
    18: "Table cells merged",
}

# Patterns for non-semantic revisions (don't affect visible content)
# These are typically locale/language settings that change based on system configuration
import re
NON_SEMANTIC_PATTERNS = [
    # Language property changes with country: "Formatted: English (Canada)", "Formatted: German (Austria)"
    # Must have the parenthesized country to distinguish from alignment like "Formatted: Left"
    re.compile(r'^Formatted:\s+[A-Z][a-z]+\s+\([A-Z][a-z,\s]+\)$'),
    # NOTE: Shading patterns are NOT included here because word_revision's FormatDescription
    # doesn't provide enough info to distinguish semantic equivalence from real failures.
    # For example, "Formatted: Pattern: 15% (Auto Foreground)" could be:
    # - Semantic equivalent: pred=pct15|auto|FFFFFF vs gold=pct15|000000|FFFFFF (OK)
    # - Real failure: pred=pct15|auto|FFFFFF vs gold=None (FAIL - shading wasn't removed)
]

def _is_empty_shape_placeholder_revision(revision_info):
    """
    Word CompareDocuments often emits delete+insert pairs for floating objects even
    when object-level properties match. These show up as empty or "/" revisions
    anchored to a shape range.
    """
    rev_type = revision_info.get('type', '')
    if rev_type not in {'Insertion', 'Deletion'}:
        return False

    range_info = revision_info.get('range') or {}
    if not range_info.get('shapes') or range_info.get('tables'):
        return False

    text = (revision_info.get('text') or '').replace('\r', '').replace('\n', '').strip()
    return text in {'', '/'}


def _is_dropcap_layout_revision(revision_info):
    """
    Drop caps serialize through a frame paragraph whose exact font size / lowered
    offset / line-spacing values can differ even when DropCap.Position and the
    rendered layout match. Treat those first-letter-only revisions as non-semantic.
    """
    rev_type = revision_info.get('type', '')
    if rev_type not in {'Property changed', 'Paragraph property changed'}:
        return False

    location = revision_info.get('location', '')
    if location != 'Main body':
        return False

    text = (revision_info.get('text') or '').replace('\r', '').replace('\n', '').strip()
    if len(text) != 1 or not text.isalpha():
        return False

    format_desc = revision_info.get('format_description', '') or ''
    dropcap_signals = (
        'Lowered by',
        'Keep with next',
        'Font Alignment: Baseline',
        'Line spacing:  Exactly',
    )
    return any(signal in format_desc for signal in dropcap_signals)


def is_semantic_revision(revision_info):
    """
    Determine if a revision is semantic (affects visible content) or non-semantic.

    Non-semantic revisions include:
    - Language property changes (e.g., "Formatted: English (Canada)")
    - Header/Footer default paragraph formatting from Normal.dotm template differences

    These don't affect visible document content but are detected by Word CompareDocuments.

    Args:
        revision_info: Dict with revision metadata (type, format_description, text, etc.)

    Returns:
        bool: True if the revision is semantic (should count), False if non-semantic
    """
    format_desc = revision_info.get('format_description', '') or ''

    # Check if this matches any non-semantic pattern
    for pattern in NON_SEMANTIC_PATTERNS:
        if pattern.match(format_desc):
            return False

    # Header/Footer default paragraph formatting differences (Normal.dotm template)
    # When Word saves a document, it auto-generates empty header/footer parts using
    # the current machine's Normal.dotm template. Different installations have different
    # default styles for "Header"/"Footer" paragraphs, causing CompareDocuments to detect
    # "Paragraph property changed" on empty paragraphs ("\r") that have no visible content.
    rev_type = revision_info.get('type', '')
    location = revision_info.get('location', '')
    text = revision_info.get('text', '')
    if rev_type == 'Paragraph property changed' and \
       ('Header' in location or 'Footer' in location) and \
       text in ('\r', '\r\n', '', None):
        return False

    if _is_dropcap_layout_revision(revision_info):
        return False

    return True

def extract_revision_info(revision):
    """Extract metadata from a single revision object."""
    try:
        rev_type = REVISION_TYPES.get(revision.Type, f"Unknown({revision.Type})")
    except:
        rev_type = "Unknown"

    try:
        format_desc = revision.FormatDescription
    except:
        format_desc = None

    try:
        text = revision.Range.Text
        # Limit text length for readability
        if len(text) > 100:
            text = text[:100] + "..."
    except:
        text = None

    try:
        range_info = {
            "start": revision.Range.Start,
            "end": revision.Range.End,
            "tables": revision.Range.Tables.Count,
            "shapes": revision.Range.InlineShapes.Count + revision.Range.ShapeRange.Count,
        }
    except:
        range_info = None

    return {
        "type": rev_type,
        "format_description": format_desc,
        "text": text,
        "range": range_info
    }

def compare_documents(word_app, doc_in_path, gold_path, compare_config):
    """
    Run Word CompareDocuments and extract all revision info.

    Args:
        word_app: Existing Word application instance
        doc_in_path: Path to input document
        gold_path: Path to gold document
        compare_config: Compare settings dict

    Returns:
        dict with revision_count, revisions list, and metadata
    """
    doc_in_path = os.path.abspath(doc_in_path)
    gold_path = os.path.abspath(gold_path)

    result = {
        "success": False,
        "error": None,
        "revision_count": 0,
        "revisions": [],
    }

    doc_in = None
    gold = None
    compare_doc = None

    try:
        # Compare the submitted documents, never a repaired in-memory version.
        # Explicit read-only/hidden options also avoid edit-lock and conversion UI.
        open_options = {
            'ReadOnly': True,
            'OpenAndRepair': False,
            'AddToRecentFiles': False,
            'Visible': False,
            'ConfirmConversions': False,
        }
        doc_in = open_normal_document(word_app, doc_in_path, **open_options)
        time.sleep(0.5)
        gold = open_normal_document(word_app, gold_path, **open_options)
        time.sleep(0.3)

        # Compare: doc_in (original) vs gold (revised)
        # This shows what changed from doc_in to gold
        compare_doc = word_app.CompareDocuments(
            OriginalDocument=doc_in,
            RevisedDocument=gold,
            Destination=2,  # wdCompareDestinationNew
            Granularity=compare_config.get("granularity", 1),  # 1=word level
            CompareFormatting=compare_config.get("compare_formatting", True),
            CompareCaseChanges=compare_config.get("compare_case_changes", True),
            CompareWhitespace=compare_config.get("compare_whitespace", True),
            CompareTables=compare_config.get("compare_tables", True),
            CompareHeaders=compare_config.get("compare_headers", True),
            CompareFootnotes=compare_config.get("compare_footnotes", True),
            CompareTextboxes=compare_config.get("compare_textboxes", True),
            CompareFields=compare_config.get("compare_fields", True),
            CompareComments=compare_config.get("compare_comments", True),
            CompareMoves=compare_config.get("compare_moves", True),
            IgnoreAllComparisonWarnings=True,
        )
        time.sleep(0.3)

        # Collect revisions from ALL locations
        all_revisions = []
        iteration_warning = None  # Track if we had issues iterating
        iteration_warnings = []
        unread_by_location = Counter()

        # 1. Main document body
        # Note: Word's COM API can fail when iterating revisions even if Count > 0
        # This happens with certain document types (e.g., shape-only documents)
        main_body_rev_count = compare_doc.Revisions.Count
        try:
            for i in range(1, main_body_rev_count + 1):
                rev = compare_doc.Revisions(i)
                rev_info = extract_revision_info(rev)
                rev_info["location"] = "Main body"
                all_revisions.append(rev_info)
        except Exception as e:
            # If iteration fails, we still trust the Count but can't get details
            iteration_warning = f"Could not iterate {main_body_rev_count} main body revisions: {e}"
            iteration_warnings.append(iteration_warning)
            unread_by_location['Main body'] += main_body_rev_count - len(all_revisions)
            print(f"      Warning: {iteration_warning}")

        def collect_story_revisions(story, location):
            # Header watermark revisions can have a valid Count while Item(i)
            # raises. Preserve the known differences, as for the main body.
            count = story.Range.Revisions.Count
            captured = 0
            try:
                for index in range(1, count + 1):
                    rev_info = extract_revision_info(story.Range.Revisions(index))
                    rev_info['location'] = location
                    all_revisions.append(rev_info)
                    captured += 1
            except Exception as exc:
                unread_by_location[location.split('(')[0].strip()] += count - captured
                iteration_warnings.append(f'Could not iterate {location} revisions: {exc}')

        # 2. Headers and footers (per section)
        for section_idx, section in enumerate(compare_doc.Sections, 1):
            # Check headers
            for header_index in [1, 2, 3]:  # Primary, First Page, Even Pages
                header = section.Headers(header_index)
                if header.Exists:
                    collect_story_revisions(header, f"Header (section {section_idx}, type {header_index})")

            # Check footers
            for footer_index in [1, 2, 3]:  # Primary, First Page, Even Pages
                footer = section.Footers(footer_index)
                if footer.Exists:
                    collect_story_revisions(footer, f"Footer (section {section_idx}, type {footer_index})")

        # 3. Footnotes
        try:
            for i in range(1, compare_doc.Footnotes.Count + 1):
                footnote = compare_doc.Footnotes(i)
                if footnote.Range.Revisions.Count > 0:
                    for j in range(1, footnote.Range.Revisions.Count + 1):
                        rev = footnote.Range.Revisions(j)
                        rev_info = extract_revision_info(rev)
                        rev_info["location"] = f"Footnote {i}"
                        all_revisions.append(rev_info)
        except:
            pass

        # 4. Endnotes
        try:
            for i in range(1, compare_doc.Endnotes.Count + 1):
                endnote = compare_doc.Endnotes(i)
                if endnote.Range.Revisions.Count > 0:
                    for j in range(1, endnote.Range.Revisions.Count + 1):
                        rev = endnote.Range.Revisions(j)
                        rev_info = extract_revision_info(rev)
                        rev_info["location"] = f"Endnote {i}"
                        all_revisions.append(rev_info)
        except:
            pass

        # 5. Text boxes (shapes with text)
        try:
            for i in range(1, compare_doc.Shapes.Count + 1):
                shape = compare_doc.Shapes(i)
                if shape.TextFrame.HasText:
                    if shape.TextFrame.TextRange.Revisions.Count > 0:
                        for j in range(1, shape.TextFrame.TextRange.Revisions.Count + 1):
                            rev = shape.TextFrame.TextRange.Revisions(j)
                            rev_info = extract_revision_info(rev)
                            rev_info["location"] = f"Text box (shape {i})"
                            all_revisions.append(rev_info)
        except:
            pass

        # 6. Comments (if any)
        try:
            for i in range(1, compare_doc.Comments.Count + 1):
                comment = compare_doc.Comments(i)
                if comment.Range.Revisions.Count > 0:
                    for j in range(1, comment.Range.Revisions.Count + 1):
                        rev = comment.Range.Revisions(j)
                        rev_info = extract_revision_info(rev)
                        rev_info["location"] = f"Comment {i}"
                        all_revisions.append(rev_info)
        except:
            pass

        # Summary
        # Calculate total revision count
        # If iteration failed, use the Count property we got before iteration
        unread_revision_count = sum(unread_by_location.values())
        total_revision_count = len(all_revisions) + unread_revision_count
        print(f"      Total revisions: {total_revision_count}")

        # Check for non-semantic revisions (e.g., language property changes)
        # Only add semantic_revision_count if it differs from total (some non-semantic exist)
        non_semantic_count = 0
        non_semantic_descriptions = []
        for rev in all_revisions:
            if not is_semantic_revision(rev):
                non_semantic_count += 1
                non_semantic_descriptions.append(rev.get("format_description", ""))

        # If there are non-semantic revisions, print a note and store semantic count
        if non_semantic_count > 0:
            semantic_revision_count = total_revision_count - non_semantic_count
            print(f"      Semantic revisions: {semantic_revision_count} (filtered {non_semantic_count} non-semantic: {set(non_semantic_descriptions)})")
            # Will add semantic_revision_count to result below
        else:
            semantic_revision_count = None  # Not needed, all revisions are semantic

        # Count by location for reporting
        location_counts = {}
        for rev in all_revisions:
            loc = rev["location"].split("(")[0].strip()  # e.g., "Header" from "Header (section 1, type 1)"
            location_counts[loc] = location_counts.get(loc, 0) + 1

        for location, count in unread_by_location.items():
            location_counts[location] = location_counts.get(location, 0) + count

        if len(location_counts) > 1 or (len(location_counts) == 1 and "Main body" not in location_counts):
            print(f"      Locations: {location_counts}")

        result["revision_count"] = total_revision_count
        result["revision_locations"] = location_counts

        # Only add semantic_revision_count if it differs from revision_count
        # (i.e., some non-semantic revisions were filtered)
        if semantic_revision_count is not None:
            result["semantic_revision_count"] = semantic_revision_count

        # Sample revisions (limit to avoid huge outputs)
        max_revisions = 50
        result["revisions"] = all_revisions[:max_revisions]

        if total_revision_count > max_revisions:
            result["revisions_truncated"] = True
            result["revisions_shown"] = max_revisions

        # Add iteration warning if present
        if iteration_warnings:
            result["iteration_warning"] = '; '.join(iteration_warnings)
            result["unread_revision_count"] = unread_revision_count

        result["success"] = True

    except Exception as e:
        result["error"] = str(e)
        result["traceback"] = traceback.format_exc()
        if isinstance(e, NormalOpenError):
            result["open_failure_path"] = e.docx_path
            result["failure_stage"] = "normal_open"

    finally:
        # Close documents in reverse order with aggressive cleanup
        docs_to_close = [(compare_doc, "compare_doc"), (gold, "gold"), (doc_in, "doc_in")]

        for doc, name in docs_to_close:
            if doc:
                try:
                    doc.Close(SaveChanges=False)
                    time.sleep(0.1)
                except Exception as e:
                    print(f"    Warning: Failed to close {name}: {e}")

    return result

# =============================================================================
# VerificationMethod Wrapper for evaluation module integration
# =============================================================================

from typing import Optional
from .base import VerificationMethod, VerificationResult
from . import register_method


@register_method
class WordRevisionMethod(WordSession, VerificationMethod):
    """
    Word revision verification method.

    Uses Word's CompareDocuments API to generate a comparison document,
    then counts revisions (tracked changes) to detect differences.
    This is a black-box approach where Word determines what differs.
    """

    def __init__(self):
        super().__init__()

    @property
    def name(self) -> str:
        return "word_revision"

    @property
    def requires_word(self) -> bool:
        return True

    def _is_case_specific_non_semantic_revision(
        self,
        revision_info: dict,
        spec: Optional[dict],
    ) -> bool:
        """
        CompareDocuments emits shape placeholder insert/delete pairs for some
        layout-coupled floating-shape cases even when the rendered output and
        object properties match. Keep this filter narrowly scoped so it does not
        hide real chart/object failures.
        """
        if not _is_empty_shape_placeholder_revision(revision_info):
            return False

        composition = (spec or {}).get('composition') or {}
        return (
            composition.get('type') == 'layout_coupled' and
            composition.get('target_object_type') == 'shape'
        )

    def verify(
        self,
        pred_path: Path,
        gold_path: Path,
        doc_in_path: Optional[Path] = None,
        spec: Optional[dict] = None
    ) -> VerificationResult:
        """Compare predicted document against gold standard using Word's CompareDocuments."""
        compare_config = {
            "granularity": 1,
            "compare_formatting": True,
            "compare_case_changes": True,
            "compare_whitespace": True,
            "compare_tables": True,
            "compare_headers": True,
            "compare_footnotes": True,
            "compare_textboxes": True,
            "compare_fields": True,
            "compare_comments": True,
            "compare_moves": True,
        }

        try:
            word = self._get_word()

            result = compare_documents(
                word,
                str(pred_path),
                str(gold_path),
                compare_config
            )

            if not result.get('success', False):
                error = result.get('error', 'Unknown error')
                if result.get('failure_stage') == 'normal_open':
                    # Close the failed comparison before independent diagnostics.
                    # Both methods use the same healthy-Word/repair protocol.
                    if not self._cleanup_word():
                        return VerificationResult.error_result('Word cleanup failed before open diagnosis')
                    gate = normal_open_gate(pred_path, gold_path)
                    if not gate['passed']:
                        return VerificationResult.from_legacy(gate)
                    return VerificationResult(
                        passed=False, message=f'ERROR: {error}', error=error,
                        details={**gate['details'], 'failure_stage': 'normal_open',
                                 'open_failure_owner': 'environment_or_unknown'},
                    )
                return VerificationResult.error_result(error)

            total_revisions = result.get('revision_count', 0)
            semantic_revisions = result.get('semantic_revision_count', total_revisions)
            revisions = result.get('revisions', [])
            diagnostics = {
                key: result[key]
                for key in ('iteration_warning', 'unread_revision_count')
                if key in result
            }

            if revisions and len(revisions) == total_revisions:
                semantic_revisions = sum(
                    1
                    for revision in revisions
                    if is_semantic_revision(revision) and
                    not self._is_case_specific_non_semantic_revision(revision, spec)
                )

            if semantic_revisions == 0:
                return VerificationResult(
                    passed=True,
                    message=f'PASS: No revisions (documents match)',
                    details={
                        'revision_count': total_revisions,
                        'semantic_revision_count': semantic_revisions,
                        **diagnostics,
                    }
                )
            else:
                return VerificationResult(
                    passed=False,
                    message=f'FAIL: {semantic_revisions} revision(s) detected',
                    details={
                        'revision_count': total_revisions,
                        'semantic_revision_count': semantic_revisions,
                        'sample_revisions': revisions[:10],
                        **diagnostics,
                    }
                )

        except Exception as e:
            return VerificationResult.error_result(str(e))

    def cleanup(self):
        """Clean up Word application."""
        self._cleanup_word()

    def __del__(self):
        """Cleanup on destruction."""
        self.cleanup()
