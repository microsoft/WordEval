"""
Test case generator for content_operations × specialized_actionable.

Generates test cases for content operations from specialized tabs (Review, Design, Draw, etc.).

Directly verifiable commands:
- InsertNewComment: Insert comment on selected text
- ReviewDeleteComment, ReviewDeleteAllCommentsInDocument, ReviewDeleteAllCommentsShown: Delete comments
- DeleteAllResolvedComments: Delete resolved comments
- AcceptAllChangesInDocAndStopTracking, RejectAllChangesInDocAndStopTracking: Track changes
- InsertInkDrawingCanvas: Insert ink drawing canvas
- WatermarkRemove: Remove watermark -> misclassification, coreactionable
"""

import json
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

from ..config import (
    COMMAND_SETS_DIR, BENCHMARK_OUTPUT_DIR,
    NAMESPACES, ONE_WAY_VARIATIONS, TOGGLE_VARIATIONS
)
from ..execution.setup import ensure_word, finalize_word
from ..execution.execution import DirectExecute
from generator.test_sets.spec_utils import write_spec_if_changed
from generator.docx_package import write_docx_package


# =============================================================================
# PRECONDITION CONFIGURATION
# =============================================================================

# Precondition type mapping (determines which input template to use)
# Default: 'default' - simple text
# Key format: 'idMso' or 'idMso:variation' for variation-specific preconditions
PRECONDITION_TYPE_MAP = {
    # Comment commands - single comment
    'ReviewDeleteComment': 'with_comment',

    # Comment commands - variation-specific
    'InsertNewComment:reply': 'with_comment',  # reply needs existing comment
    'InsertInkDrawingCanvas':'with_comment',

    # Comment commands - multiple comments (normal, with reply, resolved)
    'ReviewDeleteAllCommentsInDocument': 'with_multi_comment',
    'ReviewDeleteAllCommentsShown': 'with_multi_comment',
    'DeleteAllResolvedComments': 'with_multi_comment',
    'ReviewShowComments': 'with_multi_comment',
    'ShowCommentsPane': 'with_multi_comment',

    # Track changes - enable tracking (needs plain doc, no tracking yet)
    'ReviewTrackChanges': 'default',

    # Track changes commands need existing tracked changes
    'AcceptAllChangesInDocAndStopTracking': 'with_track_changes',
    'RejectAllChangesInDocAndStopTracking': 'with_track_changes',
    'ReviewShowInsertionsAndDeletions': 'with_track_changes',
    'ReviewShowOnlyCommentsAndFormattingInBaloons': 'with_track_changes',

    # Watermark removal needs existing watermark
    'WatermarkRemove': 'with_watermark',

    # Mail merge commands need mail merge setup
    'MailMergeHighlightMergeFields': 'with_mail_merge',
    'MailMergeInsertBarcode': 'with_mail_merge',
    'MailMergeMergeFieldInsert': 'with_mail_merge',
}


def get_precondition_type(idMso: str, variation: str = 'apply') -> str:
    """
    Determine the precondition type for a command and variation.

    Args:
        idMso: Command identifier
        variation: Test variation name ('apply', 'reply', 'on', 'off', etc.)

    Returns:
        One of: 'default', 'with_comment', 'with_multi_comment', 'with_track_changes', 'with_watermark', 'with_mail_merge'
    """
    # Check variation-specific key first
    variation_key = f'{idMso}:{variation}'
    if variation_key in PRECONDITION_TYPE_MAP:
        return PRECONDITION_TYPE_MAP[variation_key]

    # Fall back to command-level key
    return PRECONDITION_TYPE_MAP.get(idMso, 'default')


# =============================================================================
# SPECIAL HANDLING COMMANDS (SKIP)
# =============================================================================

# Commands that need special handling or are not fully automatable
SPECIAL_HANDLING_COMMANDS = {
    # Duplicate of InsertNewComment (same functionality, different tab location)
    'ReviewNewComment': 'duplicate_of_InsertNewComment',
    'ReviewDeleteAllCommentsShown':'not exectable',

    # View toggles - not verifiable in XML
    'ReviewShowComments': 'view_toggle_not_verifiable_in_xml',
    'ShowCommentsPane': 'view_toggle_not_verifiable_in_xml',
    'ShowCommentGutter': 'view_toggle_not_verifiable_in_xml',
    'ShowCommentsContextualView': 'view_toggle_not_verifiable_in_xml',
    'ReviewShowInsertionsAndDeletions': 'view_toggle_not_verifiable_in_xml',
    'ReviewShowOnlyCommentsAndFormattingInBaloons': 'view_toggle_not_verifiable_in_xml',

    # Requires ink input device
    'ReviewInkCommentNew': 'requires_ink_input_device',
    'ReviewInkCommentPen': 'requires_ink_input_device',
    'ReviewInkCommentEraser': 'requires_ink_input_device',
    'InkDeleteAll': 'requires_ink_annotations',
    'InsertInkDrawingCanvas':'requires_ink_annotations', 

    # Requires mail merge data source
    'MailMergeHighlightMergeFields': 'requires_mail_merge_data_source',
    'MailMergeInsertBarcode': 'requires_mail_merge_data_source',
    'MailMergeMergeFieldInsert': 'requires_mail_merge_data_source',

    # Requires blog post mode
    'BlogCategoryInsert': 'requires_blog_post_mode',
    'BlogInsertCategories': 'requires_blog_post_mode',
}


# =============================================================================
# COMMAND VERIFICATION PATTERNS
# =============================================================================

# Semantic verification methods for commands (for special verification logic)
# These use dedicated verification classes instead of standard xpath check
SEMANTIC_VERIFICATION_METHODS = {
    'InsertNewComment': {
        'method': 'comment_text_verification',
        'params': {'expected_text': 'Test comment.'},
        'description': 'Word COM comment text verification',
    },
    'ReviewDeleteComment': {
        'method': 'comment_verification',
        'description': 'Word COM comment verification',
    },
    'ReviewDeleteAllCommentsInDocument': {
        'method': 'comment_verification',
        'description': 'Word COM comment verification',
    },
    'DeleteAllResolvedComments': {
        'method': 'resolved_comment_verification',
        'description': 'Word COM comment resolve verification',
    },
    'AcceptAllChangesInDocAndStopTracking': {
        'method': 'revision_verification',
        'description': 'Word COM revision verification',
    },
    'RejectAllChangesInDocAndStopTracking': {
        'method': 'revision_verification',
        'description': 'Word COM revision verification',
    },
    'ReviewTrackChanges': {
        'method': 'track_changes_verification',
        'description': 'Word COM track changes verification',
    },
    'WatermarkRemove': {
        'method': 'header_footer_content_check',
        'params': {'part_type': 'header', 'condition': 'no_shape'},
        'description': 'Watermark shape removed from header',
    },
}

VARIATION_SEMANTIC_VERIFICATION_METHODS = {
    ('InsertNewComment', 'reply'): {
        'method': 'comment_reply_verification',
        'params': {
            'expected_text': 'This is a reply.',
            'expected_author': 'user2',
        },
        'description': 'Word COM reply comment verification',
    },
}


def get_semantic_verification(idMso: str, variation_name: str) -> Optional[Dict]:
    """Get semantic verification config for a command/variation."""
    variation_key = (idMso, variation_name)
    if variation_key in VARIATION_SEMANTIC_VERIFICATION_METHODS:
        return VARIATION_SEMANTIC_VERIFICATION_METHODS[variation_key]
    return SEMANTIC_VERIFICATION_METHODS.get(idMso)


# =============================================================================
# PRECONDITION SETUP FUNCTIONS
# =============================================================================

# Default input configuration
DEFAULT_INPUT = {
    'full_text': 'Alpha TARGET Omega.',
    'select_text': 'TARGET',
}

COMMENT_AUTHOR = "user1"
REPLY_AUTHOR = "user2"


def setup_precondition(doc, wd, precondition_type: str):
    """
    Setup document based on precondition type.

    Args:
        doc: Word document object
        wd: Word application object
        precondition_type: Type of precondition to setup
    """
    if precondition_type == 'default':
        # Simple text
        doc.Content.Text = DEFAULT_INPUT['full_text']

    elif precondition_type == 'with_comment':
        # Document with a single comment on 'TARGET' text
        doc.Content.Text = DEFAULT_INPUT['full_text']

        doc.Background.Fill.Visible = True 
        doc.Background.Fill.Solid()                                                                                                                       
        doc.Background.Fill.ForeColor.RGB = 15790320  # 연한 회색          

        wd.Selection.HomeKey(Unit=6)
        wd.Selection.Find.ClearFormatting()
        if wd.Selection.Find.Execute(FindText='TARGET', Forward=True):
            comment = doc.Comments.Add(wd.Selection.Range, 'This is a test comment.')
            comment.Author = COMMENT_AUTHOR

    elif precondition_type == 'with_multi_comment':
        # Document with 3 comments:
        # 1. Normal comment on "Alpha"
        # 2. Comment with reply on "TARGET"
        # 3. Resolved comment on "Omega"
        doc.Content.Text = DEFAULT_INPUT['full_text']

        # Comment 1: Normal comment on "Alpha"
        wd.Selection.HomeKey(Unit=6)
        wd.Selection.Find.ClearFormatting()
        if wd.Selection.Find.Execute(FindText='Alpha', Forward=True):
            comment1 = doc.Comments.Add(wd.Selection.Range, 'This is a normal comment.')
            comment1.Author = COMMENT_AUTHOR

        # Comment 2: Comment with reply on "TARGET"
        wd.Selection.HomeKey(Unit=6)
        wd.Selection.Find.ClearFormatting()
        if wd.Selection.Find.Execute(FindText='TARGET', Forward=True):
            comment2 = doc.Comments.Add(wd.Selection.Range, 'This is a comment with reply.')
            comment2.Author = COMMENT_AUTHOR
            # Add reply using Replies.Add(Range, Text) - Range is the comment's range
            try:
                reply = comment2.Replies.Add(comment2.Range, 'This is a reply.')
                reply.Author = REPLY_AUTHOR
            except:
                pass  # Reply may not be supported in all Word versions

        # Comment 3: Resolved comment on "Omega"
        wd.Selection.HomeKey(Unit=6)
        wd.Selection.Find.ClearFormatting()
        if wd.Selection.Find.Execute(FindText='Omega', Forward=True):
            comment3 = doc.Comments.Add(wd.Selection.Range, 'This is a resolved comment.')
            comment3.Author = COMMENT_AUTHOR
            comment3.Done = True

    elif precondition_type == 'with_track_changes':
        # Document with tracked changes (author will be fixed in post-processing)
        # Write base text with tracking off
        doc.TrackRevisions = False
        doc.Content.Text = 'Alpha TARGET Omega.'

        # Enable tracking and make change
        doc.TrackRevisions = True
        wd.Selection.HomeKey(Unit=6)  # Go to start
        wd.Selection.Find.ClearFormatting()
        wd.Selection.Find.Execute(
            FindText='TARGET',
            MatchCase=False,
            MatchWholeWord=False,
            Forward=True
        )

        # Delete and replace to create tracked change
        wd.Selection.TypeText('MODIFIED')



    elif precondition_type == 'with_watermark':
        # Document with a watermark (using BuildingBlock for WatermarkRemove ExecuteMso compatibility)
        #doc.Content.Text = DEFAULT_INPUT['full_text']

        try:
            # Load building blocks and insert watermark via BuildingBlockTypes
            wd.Templates.LoadBuildingBlocks()
            time.sleep(0.3)

            # Access watermark via BuildingBlockTypes (Index 8 = Watermarks)
            template = wd.Application.Templates(1)
            bb = template.BuildingBlockTypes(8).Categories("Disclaimers").BuildingBlocks("DRAFT 1")
            section = wd.ActiveDocument.Sections(1)
            header_range = section.Headers(1).Range  
            # wdHeaderFooterPrimary = 1
            bb.Insert(header_range, True)
            # bb.Insert(doc.Content, True)
            print(f"    Inserted watermark: DRAFT 1")

        except Exception as e:
            print(f"Warning: Could not add watermark: {e}")

    elif precondition_type == 'with_mail_merge':
        # Mail merge document (data source not set up)
        doc.Content.Text = DEFAULT_INPUT['full_text']
        print(f"    Note: Mail merge data source not set up")


def select_for_command(doc, wd, precondition_type: str, idMso: str) -> bool:
    """
    Select appropriate content based on precondition type and command.

    Args:
        doc: Word document object
        wd: Word application object
        precondition_type: Type of precondition
        idMso: Command identifier

    Returns:
        True if selection successful
    """
    # Commands that need text selection
    text_selection_commands = {
        'InsertNewComment',
        'ReviewNewComment',
        'ReviewDeleteComment',
    }

    # Commands that need cursor position after text
    cursor_after_commands = {
        'InsertInkDrawingCanvas',
    }

    # Commands that don't need selection
    no_selection_commands = {
        'ReviewDeleteAllCommentsInDocument',
        'ReviewDeleteAllCommentsShown',
        'DeleteAllResolvedComments',
        'AcceptAllChangesInDocAndStopTracking',
        'RejectAllChangesInDocAndStopTracking',
        'WatermarkRemove',
    }

    if idMso in text_selection_commands:
        # Select TARGET text
        wd.Selection.HomeKey(Unit=6)  # wdStory
        wd.Selection.Find.ClearFormatting()
        found = wd.Selection.Find.Execute(
            FindText=DEFAULT_INPUT['select_text'],
            MatchCase=False,
            MatchWholeWord=False,
            Forward=True
        )
        return found

    elif idMso in cursor_after_commands:
        # Position cursor after TARGET text
        wd.Selection.HomeKey(Unit=6)  # wdStory
        wd.Selection.Find.ClearFormatting()
        found = wd.Selection.Find.Execute(
            FindText=DEFAULT_INPUT['select_text'],
            MatchCase=False,
            MatchWholeWord=False,
            Forward=True
        )
        if found:
            wd.Selection.Collapse(Direction=0)  # wdCollapseEnd
        return found

    elif idMso in no_selection_commands:
        # No selection needed
        return True

    # Default: no selection needed
    return True


def apply_command(doc, wd, idMso: str, variation: str = 'apply') -> bool:
    """
    Apply command using COM API for commands that can't use ExecuteMso.

    Args:
        doc: Word document object
        wd: Word application object
        idMso: Command identifier
        variation: Test variation ('apply', 'reply', etc.)

    Returns:
        True if command applied successfully, None to use default ExecuteMso
    """
    try:
        if idMso == 'InsertNewComment':
            if variation == 'apply':
                # ExecuteMso opens Comment Card UI but doesn't complete
                # Use COM API: doc.Comments.Add(Range, Text)
                comment = doc.Comments.Add(wd.Selection.Range, 'Test comment.')
                comment.Author = COMMENT_AUTHOR
                time.sleep(0.5)
                return True
            elif variation == 'reply':
                # Add reply to existing comment (requires with_comment precondition)
                # Use COM API: comment.Replies.Add(Range, Text)
                if doc.Comments.Count > 0:
                    comment = doc.Comments(1)
                    reply = comment.Replies.Add(comment.Range, 'This is a reply.')
                    reply.Author = REPLY_AUTHOR
                    time.sleep(0.5)
                    return True
                return False

        if idMso == 'ReviewTrackChanges':
            # Enable track changes and make a small edit to verify
            doc.TrackRevisions = True
            # Add text at the end so we have a tracked insertion
            end_range = doc.Content
            end_range.InsertAfter(' TRACKED_EDIT')
            time.sleep(0.5)
            return True

        # Default: use ExecuteMso
        return None  # Signal to use default execution

    except Exception as e:
        print(f"    [ERROR] apply_command failed: {e}")
        return False


# =============================================================================
# GENERATOR CLASS
# =============================================================================

class ContentOperationsSpecializedActionableGenerator:
    """
    Generator for content_operations × specialized_actionable test cases.

    Handles commands from specialized tabs (Review, Mailings, Blog, Draw, Design)
    that perform content operations.

    Test set: 25 commands total
    - Directly verifiable: 10 commands
      * InsertNewComment, ReviewNewComment (insert comment)
      * ReviewDeleteComment, ReviewDeleteAllCommentsInDocument, ReviewDeleteAllCommentsShown (delete comment)
      * DeleteAllResolvedComments (delete resolved comments)
      * AcceptAllChangesInDocAndStopTracking, RejectAllChangesInDocAndStopTracking (track changes)
      * InsertInkDrawingCanvas (draw)
      * WatermarkRemove (design)
    - Skipped: 15 commands (view toggles, requires special preconditions)

    Possible/related commands (idMso) in this repo (not all are in this test set):
    - Comments (actions): InsertNewComment, ReviewNewComment, ReviewEditComment, ReviewDeleteComment,
      ReviewDeleteAllCommentsShown, ReviewDeleteAllCommentsInDocument, DeleteAllResolvedComments
    - Comments (navigation/UI): ReviewPreviousComment, ReviewNextComment, ReviewShowComments, ReviewShowAllComments,
      ShowCommentsPane, ShowCommentGutter, ShowCommentsContextualView, JumpBetweenCommentCardAndAnchor
    - Track Changes (actions): ReviewTrackChanges, ReviewTrackChangesMenu, AcceptAllChangesInDocAndStopTracking,
      RejectAllChangesInDocAndStopTracking, RevisionsLockTracking, RevisionsTrackMine, RevisionsTrackEveryone
    - Track Changes (markup/balloons/UI): ReviewShowMarkupMenu, ReviewShowMarkupAreaHighlight, ReviewShowInsertionsAndDeletions,
      ReviewBalloonsMenu, ReviewShowRevisionsInBalloons, ReviewShowRevisionsInline,
      ReviewShowOnlyCommentsAndFormattingInBaloons
    """

    def __init__(self, output_dir: Path = None):
        self.execution = DirectExecute()
        self.output_dir = output_dir or (BENCHMARK_OUTPUT_DIR / 'content_operations_x_specialized_actionable')
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.inputs_dir = self.output_dir / 'inputs'
        self.inputs_dir.mkdir(parents=True, exist_ok=True)

        # Cache for input documents
        self.input_paths = {}

    # Additional commands not in commands_by_policy_id.json classification
    # (cross-classified or synthetic commands added for PM coverage)
    ADDITIONAL_COMMANDS = [
        {
            'idMso': 'ReviewTrackChanges',
            'control_type': 'toggleButton',
            'behavior': 'content_operations',
            'scope': 'specialized_actionable',
            '_note': 'Formally settings_effect; added here for track changes infrastructure',
        },
    ]

    def load_commands(self) -> List[Dict]:
        """Load commands classified as content_operations × specialized_actionable."""
        commands_file = COMMAND_SETS_DIR / 'commands_by_policy_id.json'

        if not commands_file.exists():
            print(f"Warning: {commands_file} not found.")
            return []

        with open(commands_file, 'r', encoding='utf-8') as f:
            all_commands = json.load(f)

        commands = [
            cmd for cmd in all_commands
            if cmd.get('behavior') == 'content_operations'
            and cmd.get('scope') == 'specialized_actionable'
        ]

        # Append additional commands (cross-classified / synthetic)
        commands.extend(self.ADDITIONAL_COMMANDS)
        return commands

    def generate_input_for_precondition(self, precondition_type: str) -> Path:
        """
        Generate input document for a specific precondition type.

        Args:
            precondition_type: Precondition type ('default', 'with_comment', etc.)

        Returns:
            Path to the generated input document
        """
        input_filename = f'doc_in_{precondition_type}.docx'
        input_path = self.inputs_dir / input_filename

        # Return cached path if already generated
        if precondition_type in self.input_paths:
            return self.input_paths[precondition_type]

        if input_path.exists():
            print(f"  Reusing existing input: {input_filename}")
            self.input_paths[precondition_type] = input_path
            return input_path

        print(f"  Generating input: {input_filename}")

        wd = ensure_word()
        try:
            doc = wd.Documents.Add()

            # Setup precondition
            setup_precondition(doc, wd, precondition_type)

            doc.SaveAs(str(input_path.absolute()), FileFormat=16)
            doc.Close(False)
            time.sleep(0.2)

            # Post-process: fix revision author for track_changes
            if precondition_type == 'with_track_changes':
                self._fix_revision_author(input_path, "user1")
                print(f"    Fixed revision author to: user1")

            self.input_paths[precondition_type] = input_path
            print(f"    Created: {input_filename}")
            return input_path

        except Exception as e:
            print(f"  Error creating input: {e}")
            raise

    def _fix_revision_author(self, docx_path: Path, author: str):
        """
        Fix revision author in track changes using lxml.

        Args:
            docx_path: Path to the docx file
            author: New author name to set
        """
        from zipfile import ZipFile, ZIP_DEFLATED
        from lxml import etree
        import shutil

        # Word namespace
        w_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

        # Read all files from the docx
        with ZipFile(docx_path, 'r') as zin:
            files = {name: zin.read(name) for name in zin.namelist()}

        # Parse and modify document.xml
        doc_xml = etree.fromstring(files['word/document.xml'])

        # Find all w:ins and w:del elements and change author
        for elem in doc_xml.iter(f'{{{w_ns}}}ins', f'{{{w_ns}}}del'):
            author_attr = f'{{{w_ns}}}author'
            if author_attr in elem.attrib:
                elem.attrib[author_attr] = author

        files['word/document.xml'] = etree.tostring(
            doc_xml, xml_declaration=True, encoding='UTF-8', standalone='yes'
        )

        # Write back to a temp file, then replace original
        write_docx_package(files, docx_path)

    def get_variations_for_command(self, command: Dict) -> List[Dict]:
        """
        Get appropriate test variations based on control type.

        Args:
            command: Command dict with 'control_type' field

        Returns:
            List of variation dicts
        """
        idMso = command.get('idMso', '')
        control_type = command.get('control_type', 'button')

        # InsertNewComment has special variations: apply + reply
        if idMso == 'InsertNewComment':
            return [
                {'name': 'apply'},
                {'name': 'reply'},
            ]

        # ReviewTrackChanges: only 'apply' (enable), skip 'off' variation
        if idMso == 'ReviewTrackChanges':
            return ONE_WAY_VARIATIONS

        if control_type == 'toggleButton':
            return TOGGLE_VARIATIONS
        else:
            return ONE_WAY_VARIATIONS

    def generate_on_test(self, idMso: str, variation: Dict, input_path: Path) -> Dict:
        """
        Generate 'on' or 'apply' test: execute command.

        - doc_in: input with precondition setup
        - gold: result after executing command
        """
        test_dir = self.output_dir / idMso / variation['name']
        test_dir.mkdir(parents=True, exist_ok=True)

        gold_path = test_dir / 'gold.docx'
        spec_path = test_dir / 'spec.json'
        variation_name = variation['name']
        precondition_type = get_precondition_type(idMso, variation_name)

        result = {
            'idMso': idMso,
            'variation': variation_name,
            'success': False,
            'message': '',
            'paths': {
                'doc_in': str(input_path),
                'gold': str(gold_path),
                'spec': str(spec_path),
            }
        }

        wd = ensure_word()
        try:
            # Open input, apply command, save as gold
            doc = wd.Documents.Open(str(input_path.absolute()))

            # Select appropriate content
            if not select_for_command(doc, wd, precondition_type, idMso):
                result['message'] = f"Failed to select content for: {idMso}"
                doc.Close(False)
                return result

            # Execute the command (try apply_command first, fallback to ExecuteMso)
            apply_result = apply_command(doc, wd, idMso, variation['name'])
            if apply_result is None:
                # Use default ExecuteMso
                if not self.execution.execute(wd, idMso):
                    result['message'] = f'Failed to execute command: {idMso}'
                    doc.Close(False)
                    return result
            elif not apply_result:
                result['message'] = f'Failed to apply command: {idMso}'
                doc.Close(False)
                return result

            time.sleep(0.5)

            # Save as gold
            doc.SaveAs(str(gold_path.absolute()), FileFormat=16)
            doc.Close(False)
            time.sleep(0.2)

            # Post-process: fix revision author for track changes gold
            if idMso == 'ReviewTrackChanges':
                self._fix_revision_author(gold_path, "user1")
                print(f"    Fixed gold revision author to: user1")

            # Generate spec.json
            spec = self._generate_spec(idMso, variation, input_path)
            write_spec_if_changed(spec_path, spec)

            result['success'] = True
            result['message'] = 'Test case generated successfully'

        except Exception as e:
            result['message'] = f'Error: {str(e)}'
            import traceback
            traceback.print_exc()

        return result

    def generate_off_test(self, idMso: str, variation: Dict, input_path: Path) -> Dict:
        """
        Generate 'off' test: toggle off effect.

        - doc_in: input with effect already applied
        - gold: result after toggling off (back to precondition state)
        """
        import shutil

        test_dir = self.output_dir / idMso / variation['name']
        test_dir.mkdir(parents=True, exist_ok=True)

        doc_in_path = test_dir / 'doc_in.docx'
        gold_path = test_dir / 'gold.docx'
        spec_path = test_dir / 'spec.json'
        variation_name = variation['name']
        precondition_type = get_precondition_type(idMso, variation_name)

        result = {
            'idMso': idMso,
            'variation': variation_name,
            'success': False,
            'message': '',
            'paths': {
                'doc_in': str(doc_in_path),
                'gold': str(gold_path),
                'spec': str(spec_path),
            }
        }

        wd = ensure_word()
        try:
            # Step 1: Create doc_in (with effect applied)
            doc = wd.Documents.Open(str(input_path.absolute()))

            if not select_for_command(doc, wd, precondition_type, idMso):
                result['message'] = f"Failed to select content for: {idMso}"
                doc.Close(False)
                return result

            # Execute the command (try apply_command first, fallback to ExecuteMso)
            apply_result = apply_command(doc, wd, idMso, variation['name'])
            if apply_result is None:
                if not self.execution.execute(wd, idMso):
                    result['message'] = f'Failed to execute command: {idMso}'
                    doc.Close(False)
                    return result
            elif not apply_result:
                result['message'] = f'Failed to apply command: {idMso}'
                doc.Close(False)
                return result

            doc.SaveAs(str(doc_in_path.absolute()), FileFormat=16)
            doc.Close(False)
            time.sleep(0.2)

            # Step 2: gold = precondition state (copy of original input)
            shutil.copy(input_path, gold_path)

            # Generate spec.json
            spec = self._generate_spec(idMso, variation, input_path, is_off_test=True)
            write_spec_if_changed(spec_path, spec)

            result['success'] = True
            result['message'] = 'Test case generated successfully'

        except Exception as e:
            result['message'] = f'Error: {str(e)}'
            import traceback
            traceback.print_exc()

        return result

    def _generate_spec(
        self,
        idMso: str,
        variation: Dict,
        input_path: Path,
        is_off_test: bool = False
    ) -> Dict:
        """Generate the spec.json for a test case."""

        variation_name = variation['name']
        semantic_info = get_semantic_verification(idMso, variation_name)
        precondition_type = get_precondition_type(idMso, variation_name)

        # Calculate relative path to input
        if is_off_test:
            doc_in_ref = 'doc_in.docx'  # Local file for 'off' test
        else:
            doc_in_ref = f'../../inputs/doc_in_{precondition_type}.docx'

        spec = {
            'idMso': idMso,
            'behavior': 'content_operations',
            'scope': 'specialized_actionable',
            'variation': variation_name,
            'generated_at': datetime.now().isoformat(),

            'precondition': {
                'type': precondition_type,
                'state': variation.get('precondition', 'base'),
            },

            'files': {
                'doc_in': doc_in_ref,
                'gold': 'gold.docx',
            }
        }

        # Use semantic verification if available
        if semantic_info and not is_off_test:
            primary_spec = {
                'method': semantic_info['method'],
                'description': semantic_info.get('description', ''),
            }
            # Add params directly to primary_spec (not nested)
            primary_spec.update(semantic_info.get('params', {}))

            spec['verification'] = {
                'primary': primary_spec,
            }
        else:
            # Default to Word COM structure comparison
            spec['verification'] = {
                'primary': {
                    'method': 'structure_verification',
                    'description': 'Word COM document structure comparison',
                },
            }

        return spec

    def generate_for_command(self, command: Dict) -> List[Dict]:
        """
        Generate all test variations for a single command.

        Args:
            command: Command dict with 'idMso' and 'control_type'

        Returns:
            List of generation results
        """
        idMso = command['idMso']
        variations = self.get_variations_for_command(command)

        # Check if command should be skipped (before generating input)
        if idMso in SPECIAL_HANDLING_COMMANDS:
            # Kill Word completely to ensure clean state for next command
            try:
                finalize_word()
            except:
                pass

            # Kill any remaining Word processes
            import subprocess
            try:
                subprocess.run(['taskkill', '/F', '/IM', 'WINWORD.EXE'],
                             capture_output=True, timeout=5)
                time.sleep(1)
            except:
                pass

            # Generate skip results without input file
            results = []
            for var in variations:
                result = {
                    'idMso': idMso,
                    'variation': var['name'],
                    'success': False,
                    'message': f"Skipped: {SPECIAL_HANDLING_COMMANDS[idMso]}",
                    'skipped': True,
                    'paths': {}
                }
                results.append(result)
                print(f"  {var['name']}: SKIP - {SPECIAL_HANDLING_COMMANDS[idMso]}")
            return results

        results = []
        for var in variations:
            # Determine input path based on precondition (from PRECONDITION_TYPE_MAP)
            precondition_type = get_precondition_type(idMso, var['name'])
            var_input_path = self.generate_input_for_precondition(precondition_type)

            if var['name'] == 'on':
                result = self.generate_on_test(idMso, var, var_input_path)
            elif var['name'] == 'off':
                result = self.generate_off_test(idMso, var, var_input_path)
            elif var['name'] == 'apply':
                result = self.generate_on_test(idMso, var, var_input_path)
            elif var['name'] == 'reply':
                result = self.generate_on_test(idMso, var, var_input_path)
            else:
                continue

            results.append(result)
            if result.get('skipped'):
                status = 'SKIP'
            elif result['success']:
                status = 'OK'
            else:
                status = 'FAIL'
            print(f"  {var['name']}: {status} - {result['message']}")

        return results

    def generate_all(self, commands: List[Dict] = None, limit: int = None) -> Dict:
        """
        Generate test cases for all commands in this test set.

        Args:
            commands: List of command dicts (loads from file if None)
            limit: Optional limit on number of commands to process

        Returns:
            Summary dict with statistics
        """
        if commands is None:
            commands = self.load_commands()

        if limit:
            commands = commands[:limit]

        print(f"\n=== Generating test cases for {len(commands)} commands ===")
        print(f"Test set: content_operations × specialized_actionable")
        print(f"Output: {self.output_dir}")
        print(f"Skipped commands: {len(SPECIAL_HANDLING_COMMANDS)}\n")

        summary = {
            'test_set': 'content_operations_x_specialized_actionable',
            'generated_at': datetime.now().isoformat(),
            'total_commands': len(commands),
            'total_test_cases': 0,
            'successful': 0,
            'failed': 0,
            'skipped': 0,
            'toggle_commands': 0,
            'button_commands': 0,
            'results': [],
        }

        # Generate tests for each command
        print("Generating test cases...")
        for cmd in commands:
            idMso = cmd['idMso']
            control_type = cmd.get('control_type', 'button')

            if control_type == 'toggleButton':
                summary['toggle_commands'] += 1
            else:
                summary['button_commands'] += 1

            print(f"\n{idMso} ({control_type})...")

            results = self.generate_for_command(cmd)

            for r in results:
                summary['total_test_cases'] += 1
                if r.get('skipped'):
                    summary['skipped'] += 1
                elif r['success']:
                    summary['successful'] += 1
                else:
                    summary['failed'] += 1
                summary['results'].append(r)

        # Save summary
        summary_path = self.output_dir / 'generation_summary.json'
        with open(summary_path, 'w', encoding='utf-8') as f:
            json.dump(summary, f, indent=2, ensure_ascii=False, sort_keys=True)

        self._print_summary(summary, summary_path)

        return summary

    def _print_summary(self, summary: Dict, summary_path: Path):
        """Print generation summary."""
        print(f"\n=== Generation Complete ===")
        print(f"\nResults\n")
        print("| Status     | Count |")
        print("|-----------|-------|")
        print(f"| Successful| {summary['successful']:<5} |")
        print(f"| Skipped   | {summary.get('skipped', 0):<5} |")
        print(f"| Failed    | {summary['failed']:<5} |")

        # Print generated test cases table (including skipped)
        successful_results = [r for r in summary['results'] if r['success']]
        skipped_results = [r for r in summary['results'] if r.get('skipped')]
        failed_results = [r for r in summary['results'] if not r['success'] and not r.get('skipped')]

        if successful_results:
            print(f"\nSuccessful Test Cases ({len(successful_results)})\n")
            for r in successful_results:
                precondition_type = get_precondition_type(r['idMso'], r['variation'])
                print(f"  {r['idMso']}/{r['variation']} ({precondition_type})")

        if skipped_results:
            print(f"\nSkipped Test Cases ({len(skipped_results)})\n")
            for r in skipped_results:
                skip_reason = SPECIAL_HANDLING_COMMANDS.get(r['idMso'], 'unknown')
                print(f"  {r['idMso']}/{r['variation']} - {skip_reason}")

        if failed_results:
            print(f"\nFailed Test Cases ({len(failed_results)})\n")
            for r in failed_results:
                print(f"  {r['idMso']}/{r['variation']} - {r['message']}")

        # Print output structure
        print(f"\nOutput Structure\n")
        print(f"{self.output_dir.name}/")

        # Get unique input files
        input_files = set()
        for r in summary['results']:
            if not r.get('skipped'):
                precondition_type = get_precondition_type(r['idMso'], r['variation'])
                input_files.add(f"doc_in_{precondition_type}.docx")

        for input_file in sorted(input_files):
            print(f"  inputs/{input_file}")

        for r in successful_results:
            idMso = r['idMso']
            variation = r['variation']
            print(f"  {idMso}/{variation}/{{gold.docx, spec.json}}")

        print(f"\nTotal commands: {summary['total_commands']}")
        print(f"  - toggleButton: {summary['toggle_commands']} (2 tests each)")
        print(f"  - button: {summary['button_commands']} (1 test each)")
        print(f"Total test cases: {summary['total_test_cases']}")
        print(f"Summary saved to: {summary_path}")


# =============================================================================
# CLI
# =============================================================================

def main():
    """CLI entry point for generating content_operations × specialized_actionable tests."""
    import argparse

    parser = argparse.ArgumentParser(
        description='Generate test cases for content_operations × specialized_actionable'
    )
    parser.add_argument(
        '--command', '-c',
        help='Generate for a specific command'
    )
    parser.add_argument(
        '--limit', '-l',
        type=int,
        help='Limit number of commands to process'
    )
    parser.add_argument(
        '--output', '-o',
        help='Output directory'
    )

    args = parser.parse_args()

    output_dir = Path(args.output) if args.output else None
    generator = ContentOperationsSpecializedActionableGenerator(output_dir=output_dir)

    if args.command:
        commands = generator.load_commands()
        cmd = next((c for c in commands if c['idMso'] == args.command), None)

        if cmd:
            print(f"Generating test cases for {args.command}...")
            results = generator.generate_for_command(cmd)
            success = sum(1 for r in results if r['success'])
            print(f"\nDone: {success}/{len(results)} successful")
        else:
            print(f"Command {args.command} not found in content_operations × specialized_actionable set.")
            print(f"Attempting generation with default configuration...")
            cmd = {'idMso': args.command, 'control_type': 'button'}
            results = generator.generate_for_command(cmd)
            success = sum(1 for r in results if r['success'])
            print(f"\nDone: {success}/{len(results)} successful")
    else:
        generator.generate_all(limit=args.limit)


if __name__ == '__main__':
    main()
