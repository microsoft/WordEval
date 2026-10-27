"""
Test case generator for object_effect × contextual_actionable.

This handles commands that apply effects to contextual objects accessed via
contextual tabs, including:
- Tables (merge, split, alignment, insert, distribute, etc.)
- Pictures (crop aspect ratios, reset, fit/fill, alt text, background removal)
- SmartArt (structural node operations: add/remove, promote/demote, reorder)
- Charts (switch row/column, reset to match style)
- Equations (professional/linear format conversion, normal text)
- Text boxes (text direction)

Note: DocInfo* commands (Author / Title / FileName) used to live here but were
reclassified as content_operations and moved to
content_operations_x_contextual_actionable.
"""

import json
import os
import shutil
import time
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

try:
    from PIL import Image  # type: ignore
except ImportError:  # pragma: no cover
    Image = None


from ..config import (
    COMMAND_SETS_DIR, COMMAND_SPECS_DIR, BENCHMARK_OUTPUT_DIR,
    NAMESPACES, TOGGLE_VARIATIONS, ONE_WAY_VARIATIONS,
)
from ..execution.setup import ensure_word, finalize_word
from ..execution.execution import DirectExecute
from generator.test_sets.spec_utils import write_spec_if_changed

# =============================================================================
# COMMAND CATEGORIZATION BY OBJECT TYPE
# =============================================================================

# Table commands - require a table with cells
TABLE_COMMANDS = {
    'TableRowsDelete', 'TableColumnsDelete', 'MergeCells', 'SplitCells',
    'TableSplitTable', 'TableColumnsDistribute', 'TableRowsDistribute',
    'TableDelete', 'TableRowsInsertAboveWord', 'TableRowsInsertBelowWord',
    'TableColumnsInsertLeft', 'TableColumnsInsertRight', 'TableAutoFitContents',
    'TableAutoFitWindow',
    # Table cell alignment (9 commands)
    'TableCellAlignTopLeft', 'TableCellAlignTopCenter', 'TableCellAlignTopRight',
    'TableCellAlignMiddleLeft', 'TableCellAlignMiddleCenter', 'TableCellAlignMiddleRight',
    'TableCellAlignBottomLeft', 'TableCellAlignBottomCenter', 'TableCellAlignBottomRight',
    'TableStyleClear',
}

# Commands that CHANGE text content (delete/add cells, merge, split)
TEXT_CHANGING_COMMANDS = {
    'TableRowsDelete', 'TableColumnsDelete', 'MergeCells', 'SplitCells', 'TableDelete',
    'TableSplitTable', 'TableRowsInsertAboveWord', 'TableRowsInsertBelowWord',
    'TableColumnsInsertLeft', 'TableColumnsInsertRight',
}

# Picture/Image commands - require a picture
PICTURE_COMMANDS = {
    'PictureReset', 'PictureResetAndSize', 'AlternativeText',
    # Crop aspect ratio commands (12 commands)
    'PictureCropAspectRatio1To1', 'PictureCropAspectRatio2To3', 'PictureCropAspectRatio3To5',
    'PictureCropAspectRatio4To5', 'PictureFitCrop', 'PictureFillCrop',
    'PictureCropAspectRatio3To2', 'PictureCropAspectRatio5To3', 'PictureCropAspectRatio5To4',
    'PictureCropAspectRatio3To4', 'PictureCropAspectRatio4To3',
    'PictureCropAspectRatio16To9', 'PictureCropAspectRatio16To10',
}

# SmartArt commands - require a SmartArt graphic
SMARTART_COMMANDS = {
    # Add shape commands
    'SmartArtAddShapeAfter', 'SmartArtAddShapeBefore',
    'SmartArtAddShapeAbove', 'SmartArtAddShapeBelow',
    'SmartArtAddAssistant',
    # Level commands
    'SmartArtPromote', 'SmartArtDemote',
    # Reorder commands
    'SmartArtReorderUp', 'SmartArtReorderDown',
    # Size commands
    'SmartArtLargerShape', 'SmartArtSmallerShape',
    # Toggle/reset
    'SmartArtRightToLeft', 'SmartArtResetGraphic',
}

# Chart commands - require a chart
CHART_COMMANDS = {'ChartSwitchRowColumn', 'ChartResetToMatchStyle'}

# WordArt/TextBox commands
WORDART_TEXTBOX_COMMANDS = {'TextDirection'}

# Equation commands - require an equation
EQUATION_COMMANDS = {
    'EquationNormalText', 'EquationProfessionalAll', 'EquationProfessionalOne',
    'EquationLinearAll', 'EquationLinearOne',
}

# All expanded commands (non-table, validated via expanded test set)
ALL_EXPANDED_COMMANDS = (PICTURE_COMMANDS | WORDART_TEXTBOX_COMMANDS
                         | EQUATION_COMMANDS | CHART_COMMANDS | SMARTART_COMMANDS)


# =============================================================================
# SPECIAL HANDLING (Non-automatable or Complex)
# =============================================================================

SPECIAL_HANDLING_COMMANDS = {
    # Pending validation - generation skipped until verifier/setup is finalized.
    # Picture
    'SmartArtAddShape': 'pending_validation',
    'SmartArtOrganizationChartLeftHanging': 'pending_validation',
    'SmartArtOrganizationChartRightHanging': 'pending_validation',
    'SmartArtOrganizationChartBoth': 'pending_validation',
    'SmartArtOrganizationChartStandard': 'pending_validation',
    'SmartArtAddBullet': 'pending_validation',
    'WordArtClear': 'pending_validation',
    'SmartArtRecolorImages': 'pending_validation',
    'ChartFormatSelection': 'pending_validation',
    'ChartRefresh': 'pending_validation',
    'HeaderFooterLinkToPrevious': 'pending_validation',
    'EquationLaTexFormat': 'pending_validation',
    'EquationUnicodeFormat': 'pending_validation',
    'EquationConvert': 'pending_validation',
}


# =============================================================================
# PRECONDITION TYPE MAPPING
# =============================================================================

def get_precondition_type(idMso: str) -> str:
    """
    Determine the precondition type (object type) for a command.

    Returns:
        One of: 'table', 'table_merged', 'picture', 'textbox', 'info',
                'equation', 'chart', 'chart_custom', 'smartart_list',
                'smartart_hierarchy', 'smartart_orgchart', 'unknown'
    """
    if idMso == 'SplitCells':
        return 'table_merged'
    elif idMso in TABLE_COMMANDS:
        return 'table'
    elif idMso in PICTURE_COMMANDS:
        return 'picture'
    elif idMso in WORDART_TEXTBOX_COMMANDS:
        return 'textbox'
    elif idMso in EQUATION_COMMANDS:
        return 'equation'
    elif idMso == 'ChartResetToMatchStyle':
        return 'chart_custom'
    elif idMso in CHART_COMMANDS:
        return 'chart'
    elif idMso in SMARTART_COMMANDS:
        if idMso in ('SmartArtAddShapeAbove', 'SmartArtAddShapeBelow',
                     'SmartArtPromote', 'SmartArtDemote'):
            return 'smartart_hierarchy'
        elif idMso == 'SmartArtAddAssistant':
            return 'smartart_orgchart'
        else:
            return 'smartart_list'
    else:
        return 'unknown'


# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def apply_table_borders(table):
    """
    Apply visible borders (single line) to all edges of a table.
    
    This makes table structure clearly visible for visual validation,
    addressing the validation report issue where borderless tables
    look like floating text and make operations hard to verify.
    
    Args:
        table: Word Table object
    """
    # wdBorderTop=1, wdBorderLeft=2, wdBorderBottom=3, wdBorderRight=4
    # wdBorderHorizontal=5, wdBorderVertical=6
    # wdLineStyleSingle=1
    try:
        for border_id in [1, 2, 3, 4, 5, 6]:  # All border types
            table.Borders(border_id).LineStyle = 1  # wdLineStyleSingle
    except:
        pass


# =============================================================================
# SETUP FUNCTIONS
# =============================================================================

def setup_precondition(doc, wd, precondition_type: str, idMso: Optional[str] = None):
    """
    Setup document based on precondition type (object type).

    Args:
        doc: Word document object
        wd: Word application object
        precondition_type: Type of object to create
        idMso: Command identifier (for special setup)
    """
    if precondition_type == 'table':
        # Insert a 3x3 table with some text
        table = doc.Tables.Add(
            Range=wd.Selection.Range,
            NumRows=3,
            NumColumns=3
        )
        # Add some text to cells
        for row in range(1, 4):
            for col in range(1, 4):
                table.Cell(row, col).Range.Text = f"R{row}C{col}"

        # Use auto-fit to CONTENTS mode so TableAutoFitWindow has a visible effect.
        # Then set explicit row heights for alignment visibility
        try:
            table.AutoFitBehavior(1)  # wdAutoFitContent - auto-fit to contents
            # Set row heights to 1 inch (72 points) - tall enough to clearly show vertical alignment
            for row in table.Rows:
                row.Height = 72  # 1 inch in points
                row.HeightRule = 1  # wdRowHeightExactly - exact height
            # Set column widths to 2 inches (144 points) - wide enough for horizontal alignment visibility
            for col in table.Columns:
                col.Width = 144  # 2 inches
        except:
            pass

        # Set cell alignment to middle-center (non-default)
        # This ensures TableCellAlignTopLeft and other alignment commands have observable effects
        # (Default is top-left, so TableCellAlignTopLeft would be a no-op otherwise)
        try:
            for row in range(1, 4):
                for col in range(1, 4):
                    cell = table.Cell(row, col)
                    cell.VerticalAlignment = 1  # wdCellAlignVerticalCenter
                    cell.Range.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter
        except:
            pass

        # Add visible borders to make table structure clear for visual validation
        apply_table_borders(table)

    elif precondition_type == 'table_merged':
        # 3x3 table with top-left 2x2 cells merged into one
        table = doc.Tables.Add(
            Range=wd.Selection.Range,
            NumRows=3,
            NumColumns=3
        )
        # Add text to all cells first
        for row in range(1, 4):
            for col in range(1, 4):
                table.Cell(row, col).Range.Text = f"R{row}C{col}"

        # Set cell dimensions for visibility
        try:
            for row in table.Rows:
                row.Height = 72  # 1 inch
                row.HeightRule = 1  # wdRowHeightExactly
            for col in table.Columns:
                col.Width = 144  # 2 inches
        except:
            pass

        # Merge the top-left 2x2 cells using Cell.Merge (more reliable than Selection-based merge)
        table.Cell(1, 1).Merge(MergeTo=table.Cell(2, 2))
        time.sleep(0.3)
        # Put text in merged cell (after merge, Cell(1,1) is the merged cell)
        table.Cell(1, 1).Range.Text = "Merged 2x2"

        # Add visible borders
        apply_table_borders(table)

    elif precondition_type == 'picture':
        # Create an actual image file and insert it

        # Create a test image with 16:9 aspect ratio and visual markers
        # This makes crop commands visually obvious (cropping to 4:3, 1:1, etc. will be different)
        img_path = Path(tempfile.gettempdir()) / 'test_image.png'
        try:
            if Image is None:
                raise ImportError("Pillow (PIL) is not installed")

            # Create a 640x360 image (16:9 aspect ratio) with gradient and border markers
            width, height = 640, 360
            img = Image.new('RGB', (width, height), color=(73, 109, 137))

            # Add visual markers to make cropping visible
            from PIL import ImageDraw
            draw = ImageDraw.Draw(img)

            # Draw colored corners to show when they get cropped
            corner_size = 60
            # Top-left: Red
            draw.rectangle([0, 0, corner_size, corner_size], fill=(255, 80, 80))
            # Top-right: Green
            draw.rectangle([width-corner_size, 0, width, corner_size], fill=(80, 200, 80))
            # Bottom-left: Blue
            draw.rectangle([0, height-corner_size, corner_size, height], fill=(80, 80, 255))
            # Bottom-right: Yellow
            draw.rectangle([width-corner_size, height-corner_size, width, height], fill=(255, 220, 80))

            # Draw center cross to show center point
            center_x, center_y = width // 2, height // 2
            draw.line([(center_x - 30, center_y), (center_x + 30, center_y)], fill=(255, 255, 255), width=3)
            draw.line([(center_x, center_y - 30), (center_x, center_y + 30)], fill=(255, 255, 255), width=3)

            # Draw border
            draw.rectangle([0, 0, width-1, height-1], outline=(255, 255, 255), width=2)

            img.save(str(img_path))

            # Insert the image as an inline shape (default for pictures)
            picture = doc.InlineShapes.AddPicture(
                FileName=str(img_path),
                LinkToFile=False,
                SaveWithDocument=True,
                Range=wd.Selection.Range
            )
            time.sleep(0.2)
        except Exception as e:
            print(f"  Warning: Failed to create picture: {e}")
            # Fallback to shape
            shape = doc.Shapes.AddShape(1, 100, 100, 200, 150)
            shape.Fill.ForeColor.RGB = 255

    elif precondition_type == 'textbox':
        # Insert a text box
        shape = doc.Shapes.AddTextbox(
            Orientation=1,  # msoTextOrientationHorizontal
            Left=100,
            Top=100,
            Width=200,
            Height=100
        )
        # Add text content
        shape.TextFrame.TextRange.Text = "Line 1\nLine 2\nLine 3"

    elif precondition_type == 'chart':
        # Insert a clustered column chart with default theme colors
        try:
            chart_shape = doc.InlineShapes.AddChart(
                Type=51,  # xlColumnClustered
                Range=wd.Selection.Range
            )
            time.sleep(1)  # Chart creation takes time
        except Exception as e:
            print(f"  Warning: chart setup error: {e}")
            shape = doc.Shapes.AddShape(1, 50, 50, 300, 200)
            shape.TextFrame.TextRange.Text = "Chart Placeholder"

    elif precondition_type == 'chart_custom':
        # Insert a clustered column chart with custom formatting (for ResetToMatchStyle)
        try:
            chart_shape = doc.InlineShapes.AddChart(
                Type=51,  # xlColumnClustered
                Range=wd.Selection.Range
            )
            time.sleep(1)  # Chart creation takes time
            chart = chart_shape.Chart
            chart.ChartColor = 3  # Custom color scheme
            # Apply direct formatting to series 1 so ResetToMatchStyle has something to reset
            chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = 0x0000FF  # Red fill (BGR)
            time.sleep(0.3)
        except Exception as e:
            print(f"  Warning: chart setup error: {e}")
            shape = doc.Shapes.AddShape(1, 50, 50, 300, 200)
            shape.TextFrame.TextRange.Text = "Chart Placeholder"

    elif precondition_type == 'equation':
        # Insert 4 equations in LPPL order (Linear, Professional, Professional, Linear)
        # Mixed state ensures every All/One command produces gold != doc_in
        eq_texts = ["x^2+y=z", "a^2+b=c", "m^2+n=p", "u^2+v=w"]
        professional_indices = [1, 2]  # 0-based: eq2 and eq3 get BuildUp()
        try:
            for idx, eq_text in enumerate(eq_texts):
                if idx == 0:
                    equation_range = doc.Range()
                    equation_range.Collapse(1)  # wdCollapseStart
                else:
                    wd.Selection.HomeKey(6)  # wdStory
                    time.sleep(0.1)
                    wd.Selection.EndKey(6)  # wdStory
                    time.sleep(0.1)
                    wd.Selection.TypeParagraph()
                    time.sleep(0.1)
                    equation_range = wd.Selection.Range

                omath = equation_range.OMaths.Add(equation_range)
                wd.Selection.TypeText(eq_text)
                time.sleep(0.3)

                if idx in professional_indices:
                    # BuildUp converts Linear -> Professional
                    doc.OMaths(idx + 1).BuildUp()
                    time.sleep(0.2)
        except Exception as e:
            print(f"  Warning: equation setup error: {e}")
            wd.Selection.TypeText("x^2+y=z")

    elif precondition_type == 'smartart_list':
        # Basic Block List with 5 labeled nodes + custom color for ResetGraphic
        try:
            layouts = wd.Application.SmartArtLayouts
            sa_shape = doc.Shapes.AddSmartArt(layouts(1), 50, 50, 400, 300)
            time.sleep(1)
            sa = sa_shape.SmartArt
            labels = ['Alpha', 'Beta', 'Charlie', 'Delta', 'Echo']
            for i in range(1, min(sa.Nodes.Count + 1, len(labels) + 1)):
                sa.Nodes(i).TextFrame2.TextRange.Text = labels[i - 1]
            # Apply custom color for ResetGraphic testing
            try:
                colors = wd.Application.SmartArtColors
                if colors.Count >= 2:
                    sa.Color = colors(2)
            except:
                pass
            time.sleep(0.3)
        except Exception as e:
            print(f"  Warning: SmartArt list setup error: {e}")
            shape = doc.Shapes.AddShape(1, 50, 50, 400, 300)
            shape.TextFrame.TextRange.Text = "SmartArt Placeholder"

    elif precondition_type == 'smartart_hierarchy':
        # Hierarchy layout with 3 levels for Promote/Demote/AddAbove/AddBelow
        try:
            layouts = wd.Application.SmartArtLayouts
            sa_shape = doc.Shapes.AddSmartArt(layouts(102), 50, 50, 400, 300)
            time.sleep(1)
            sa = sa_shape.SmartArt
            # Default hierarchy has: Root -> Child1 -> Grandchild1/Grandchild2, Child2 -> Grandchild3
            # Label all nodes via AllNodes
            labels = ['Boss', 'Team1', 'Member1', 'Member2', 'Team2', 'Member3']
            all_nodes = sa.AllNodes
            for i in range(1, min(all_nodes.Count + 1, len(labels) + 1)):
                all_nodes(i).TextFrame2.TextRange.Text = labels[i - 1]
            time.sleep(0.3)
        except Exception as e:
            print(f"  Warning: SmartArt hierarchy setup error: {e}")
            shape = doc.Shapes.AddShape(1, 50, 50, 400, 300)
            shape.TextFrame.TextRange.Text = "Hierarchy Placeholder"

    elif precondition_type == 'smartart_orgchart':
        # Organization Chart for AddAssistant
        try:
            layouts = wd.Application.SmartArtLayouts
            sa_shape = doc.Shapes.AddSmartArt(layouts(97), 50, 50, 400, 300)
            time.sleep(1)
            sa = sa_shape.SmartArt
            top = sa.Nodes(1)
            top.TextFrame2.TextRange.Text = 'Boss'
            for i in range(1, top.Nodes.Count + 1):
                top.Nodes(i).TextFrame2.TextRange.Text = f'Sub{i}'
            time.sleep(0.3)
        except Exception as e:
            print(f"  Warning: SmartArt org chart setup error: {e}")
            shape = doc.Shapes.AddShape(1, 50, 50, 400, 300)
            shape.TextFrame.TextRange.Text = "OrgChart Placeholder"

    elif precondition_type == 'header_footer':
        # Create a document with 2 sections for HeaderFooterLinkToPrevious
        # First section with header
        if doc.Sections.Count > 0:
            section1 = doc.Sections(1)
            header1 = section1.Headers(1)  # wdHeaderFooterPrimary
            header1.Range.Text = "Section 1 Header"

        # Add a section break and second section
        wd.Selection.EndKey(Unit=6)  # wdStory - go to end
        wd.Selection.InsertBreak(2)  # wdSectionBreakNextPage
        time.sleep(0.2)

        # Add content to second section
        wd.Selection.TypeText("Section 2 content")

        # Now we have 2 sections, second one can be linked/unlinked from first

    else:
        # Default: insert a simple shape
        doc.Shapes.AddShape(1, 100, 100, 150, 100)


def select_for_command(doc, wd, execution, precondition_type: str, select_text: str = "TARGET", idMso: Optional[str] = None) -> bool:
    """
    Select appropriate object(s) based on precondition type and command.

    Args:
        doc: Word document object
        wd: Word application object
        execution: Execution helper object
        precondition_type: Type of precondition/object
        select_text: Unused (kept for backward compatibility)
        idMso: Command identifier (used to determine selection strategy)

    Returns:
        True if selection successful
    """
    if precondition_type == 'table':
        if doc.Tables.Count > 0:
            table = doc.Tables(1)

            # Table-specific selection logic
            if idMso == 'TableSplitTable':
                # Split table at row 2 - cursor needs to be in row 2
                table.Cell(2, 1).Select()
            elif idMso in ['MergeCells'] or 'TableCellAlign' in idMso:
                # Select multiple cells for merge/alignment
                table.Cell(1, 1).Range.End = table.Cell(2, 2).Range.End
                wd.Selection.SetRange(table.Cell(1, 1).Range.Start, table.Cell(2, 2).Range.End)
            else:
                # Default: select first cell
                table.Cell(1, 1).Select()
            return True
        return False

    elif precondition_type == 'table_merged':
        if doc.Tables.Count > 0:
            doc.Tables(1).Cell(1, 1).Select()
            return True
        return False

    elif precondition_type == 'picture':
        # Pictures are typically InlineShapes, not Shapes
        if doc.InlineShapes.Count > 0:
            doc.InlineShapes(1).Select()
            return True
        elif doc.Shapes.Count > 0:
            doc.Shapes(1).Select()
            return True
        return False

    elif precondition_type == 'textbox':
        # Select the shape/object
        if doc.Shapes.Count > 0:
            doc.Shapes(1).Select()
            return True
        return False

    elif precondition_type in ('chart', 'chart_custom'):
        # Charts are InlineShapes
        if doc.InlineShapes.Count > 0:
            doc.InlineShapes(1).Select()
            return True
        return False

    elif precondition_type == 'equation':
        # Select appropriate equation based on command
        # LPPL layout: Linear* selects eq2(P), Professional* selects eq1(L)
        if doc.OMaths.Count > 0:
            if idMso in ('EquationLinearAll', 'EquationLinearOne'):
                eq_idx = min(2, doc.OMaths.Count)  # eq2 (1-based)
            else:
                eq_idx = 1  # eq1 for Professional* and NormalText
            doc.OMaths(eq_idx).Range.Select()
            return True
        return False

    elif precondition_type in ('smartart_list', 'smartart_hierarchy', 'smartart_orgchart'):
        # Select the SmartArt shape
        for i in range(1, doc.Shapes.Count + 1):
            try:
                if doc.Shapes(i).HasSmartArt:
                    doc.Shapes(i).Select()
                    return True
            except:
                pass
        return False

    elif precondition_type == 'header_footer':
        # For HeaderFooterLinkToPrevious, we need to be in the second section's header
        if doc.Sections.Count >= 2:
            # Enter header editing mode for section 2
            try:
                # Get second section's header
                section2 = doc.Sections(2)
                header2 = section2.Headers(1)  # wdHeaderFooterPrimary
                # Select the header range
                header2.Range.Select()
                return True
            except:
                return False
        return False

    else:
        # Default: select first shape if available
        if doc.Shapes.Count > 0:
            doc.Shapes(1).Select()
            return True
        return False


# =============================================================================
# API-BASED EXECUTION HELPERS
# =============================================================================

def _find_smartart(doc):
    """Find first SmartArt in document, return (shape, smartart) or (None, None)."""
    for i in range(1, doc.Shapes.Count + 1):
        try:
            if doc.Shapes(i).HasSmartArt:
                return doc.Shapes(i), doc.Shapes(i).SmartArt
        except:
            pass
    return None, None


def _find_node_by_text(sa, text):
    """Find a node by its text label (searches AllNodes)."""
    try:
        all_nodes = sa.AllNodes
        for i in range(1, all_nodes.Count + 1):
            if all_nodes(i).TextFrame2.TextRange.Text.strip() == text:
                return all_nodes(i)
    except:
        pass
    return None


def execute_via_api(wd, doc, idMso: str) -> bool:
    """
    Execute command using Word object model APIs instead of ExecuteMso.

    Args:
        wd: Word application object
        doc: Word document object
        idMso: Command identifier

    Returns:
        True if successful, False otherwise
    """
    try:
        # Get the selected shape/object
        selection = wd.Selection

        # Header/Footer commands
        if idMso == 'HeaderFooterLinkToPrevious':
            # Toggle LinkToPrevious for current section's header
            if doc.Sections.Count > 1:
                section = selection.Sections(1)
                header = section.Headers(1)  # wdHeaderFooterPrimary
                header.LinkToPrevious = not header.LinkToPrevious
                return True


        # Table Distribute commands
        elif idMso == 'TableColumnsDistribute':
            if doc.Tables.Count > 0:
                table = selection.Tables(1)
                table.Columns.DistributeWidth()
                return True

        elif idMso == 'TableRowsDistribute':
            if doc.Tables.Count > 0:
                table = selection.Tables(1)
                table.Rows.DistributeHeight()
                return True

        # TextDirection - rotate text orientation in textbox
        elif idMso == 'TextDirection':
            if selection.ShapeRange.Count > 0:
                shape = selection.ShapeRange(1)
                # msoTextOrientationHorizontal=1, msoTextOrientationUpward=2,
                # msoTextOrientationDownward=3, msoTextOrientationVerticalFarEast=4
                current = shape.TextFrame.Orientation
                # Cycle through: 1 -> 3 -> 2 -> 1 (horizontal -> downward -> upward -> horizontal)
                orientation_cycle = {1: 3, 3: 2, 2: 1, 4: 1, 5: 1}
                shape.TextFrame.Orientation = orientation_cycle.get(current, 3)
                return True

        # Picture crop aspect ratio commands
        elif idMso.startswith('PictureCropAspectRatio'):
            # Parse aspect ratio from command name
            aspect_ratios = {
                'PictureCropAspectRatio1To1': (1, 1),
                'PictureCropAspectRatio2To3': (2, 3),
                'PictureCropAspectRatio3To2': (3, 2),
                'PictureCropAspectRatio3To4': (3, 4),
                'PictureCropAspectRatio4To3': (4, 3),
                'PictureCropAspectRatio3To5': (3, 5),
                'PictureCropAspectRatio5To3': (5, 3),
                'PictureCropAspectRatio4To5': (4, 5),
                'PictureCropAspectRatio5To4': (5, 4),
                'PictureCropAspectRatio16To9': (16, 9),
                'PictureCropAspectRatio16To10': (16, 10),
            }
            if idMso in aspect_ratios:
                target_w, target_h = aspect_ratios[idMso]
                target_ratio = target_w / target_h

                # Get the picture (InlineShape)
                if doc.InlineShapes.Count > 0:
                    pic = doc.InlineShapes(1)
                    pic_width = pic.Width
                    pic_height = pic.Height
                    current_ratio = pic_width / pic_height

                    # Calculate crop values to achieve target aspect ratio
                    if current_ratio > target_ratio:
                        # Picture is wider than target - crop left and right
                        new_width = pic_height * target_ratio
                        crop_amount = (pic_width - new_width) / 2
                        pic.PictureFormat.CropLeft = crop_amount
                        pic.PictureFormat.CropRight = crop_amount
                    else:
                        # Picture is taller than target - crop top and bottom
                        new_height = pic_width / target_ratio
                        crop_amount = (pic_height - new_height) / 2
                        pic.PictureFormat.CropTop = crop_amount
                        pic.PictureFormat.CropBottom = crop_amount
                    return True

        # AlternativeText - set alt text on picture
        elif idMso == 'AlternativeText':
            if selection.InlineShapes.Count > 0:
                selection.InlineShapes(1).AlternativeText = 'A colorful test image with corner markers'
            elif selection.ShapeRange.Count > 0:
                selection.ShapeRange(1).AlternativeText = 'A colorful test image with corner markers'
            return True

        # PictureReset - reset picture to original
        elif idMso == 'PictureReset':
            if doc.InlineShapes.Count > 0:
                pic = doc.InlineShapes(1)
                pic.Reset()
                return True

        # PictureResetAndSize - reset picture size
        elif idMso == 'PictureResetAndSize':
            if doc.InlineShapes.Count > 0:
                pic = doc.InlineShapes(1)
                pic.Reset()
                # Also reset scaling
                pic.ScaleHeight = 100
                pic.ScaleWidth = 100
                return True

        # PictureFitCrop / PictureFillCrop - these need the picture to be in a shape container
        elif idMso in ['PictureFitCrop', 'PictureFillCrop']:
            # These commands work on pictures in shapes, not inline shapes
            # For inline shapes, we simulate by adjusting crop
            if doc.InlineShapes.Count > 0:
                pic = doc.InlineShapes(1)
                if idMso == 'PictureFitCrop':
                    # Fit: ensure entire picture visible (may have letterboxing)
                    pic.PictureFormat.CropLeft = 0
                    pic.PictureFormat.CropRight = 0
                    pic.PictureFormat.CropTop = 0
                    pic.PictureFormat.CropBottom = 0
                else:  # PictureFillCrop
                    # Fill: crop to fill container (may clip edges)
                    # Apply a small uniform crop to show "fill" effect
                    pic.PictureFormat.CropLeft = 10
                    pic.PictureFormat.CropRight = 10
                    pic.PictureFormat.CropTop = 10
                    pic.PictureFormat.CropBottom = 10
                return True

        elif idMso == 'SplitCells':
            selection.Cells.Split(NumRows=2, NumColumns=2)
            return True

        # ChartSwitchRowColumn - toggle PlotBy between xlRows(1) and xlColumns(2)
        elif idMso == 'ChartSwitchRowColumn':
            if doc.InlineShapes.Count > 0 and doc.InlineShapes(1).HasChart:
                chart = doc.InlineShapes(1).Chart
                chart.ChartData.Activate()  # Enter edit mode (required after doc reopen)
                time.sleep(0.5)
                # Toggle: xlRows=1 <-> xlColumns=2
                chart.PlotBy = 1 if chart.PlotBy == 2 else 2
                time.sleep(0.3)
                return True
            return False

        # ChartResetToMatchStyle - clear custom formatting
        elif idMso == 'ChartResetToMatchStyle':
            if doc.InlineShapes.Count > 0 and doc.InlineShapes(1).HasChart:
                chart = doc.InlineShapes(1).Chart
                chart.ChartData.Activate()  # Enter edit mode (required after doc reopen)
                time.sleep(0.5)
                chart.ClearToMatchStyle()
                time.sleep(0.3)
                return True
            return False

        # =====================================================================
        # SmartArt structural commands -- use SmartArt node API directly
        # =====================================================================

        # SmartArtAddShapeAfter - add node after target ("Charlie")
        elif idMso == 'SmartArtAddShapeAfter':
            _, sa = _find_smartart(doc)
            if sa:
                target = _find_node_by_text(sa, 'Charlie')
                if target:
                    new_node = target.AddNode(2)  # msoSmartArtNodeAfter
                    new_node.TextFrame2.TextRange.Text = 'New'
                    return True
            return False

        # SmartArtAddShapeBefore - add node before target ("Charlie")
        elif idMso == 'SmartArtAddShapeBefore':
            _, sa = _find_smartart(doc)
            if sa:
                target = _find_node_by_text(sa, 'Charlie')
                if target:
                    new_node = target.AddNode(3)  # msoSmartArtNodeBefore
                    new_node.TextFrame2.TextRange.Text = 'New'
                    return True
            return False

        # SmartArtAddShapeAbove - add node above target in hierarchy ("Member1")
        elif idMso == 'SmartArtAddShapeAbove':
            _, sa = _find_smartart(doc)
            if sa:
                target = _find_node_by_text(sa, 'Member1')
                if target:
                    new_node = target.AddNode(4)  # msoSmartArtNodeAbove
                    new_node.TextFrame2.TextRange.Text = 'New'
                    return True
            return False

        # SmartArtAddShapeBelow - add node below target in hierarchy ("Team1")
        elif idMso == 'SmartArtAddShapeBelow':
            _, sa = _find_smartart(doc)
            if sa:
                target = _find_node_by_text(sa, 'Team1')
                if target:
                    new_node = target.AddNode(5)  # msoSmartArtNodeBelow
                    new_node.TextFrame2.TextRange.Text = 'New'
                    return True
            return False

        # SmartArtAddAssistant - add child after last subordinate of "Boss"
        elif idMso == 'SmartArtAddAssistant':
            _, sa = _find_smartart(doc)
            if sa and sa.Nodes.Count > 0:
                boss = sa.Nodes(1)
                if boss.Nodes.Count > 0:
                    last_child = boss.Nodes(boss.Nodes.Count)
                    asst = last_child.AddNode(2)  # msoSmartArtNodeAfter
                else:
                    asst = boss.AddNode(5)  # msoSmartArtNodeBelow (fallback)
                asst.TextFrame2.TextRange.Text = 'Assistant'
                return True
            return False

        # SmartArtPromote - promote "Member1" (level 3 -> level 2)
        elif idMso == 'SmartArtPromote':
            _, sa = _find_smartart(doc)
            if sa:
                target = _find_node_by_text(sa, 'Member1')
                if target:
                    target.Promote()
                    return True
            return False

        # SmartArtDemote - demote "Team2" (level 2 -> level 3)
        elif idMso == 'SmartArtDemote':
            _, sa = _find_smartart(doc)
            if sa:
                target = _find_node_by_text(sa, 'Team2')
                if target:
                    target.Demote()
                    return True
            return False

        # SmartArtReorderUp - move "Charlie" (node 3) up
        elif idMso == 'SmartArtReorderUp':
            _, sa = _find_smartart(doc)
            if sa:
                target = _find_node_by_text(sa, 'Charlie')
                if target:
                    target.ReorderUp()
                    return True
            return False

        # SmartArtReorderDown - move "Charlie" (node 3) down
        elif idMso == 'SmartArtReorderDown':
            _, sa = _find_smartart(doc)
            if sa:
                target = _find_node_by_text(sa, 'Charlie')
                if target:
                    target.ReorderDown()
                    return True
            return False

        # SmartArtLargerShape - make "Charlie" larger
        elif idMso == 'SmartArtLargerShape':
            _, sa = _find_smartart(doc)
            if sa:
                target = _find_node_by_text(sa, 'Charlie')
                if target:
                    target.Larger()
                    target.Larger()  # Apply twice for more noticeable change
                    return True
            return False

        # SmartArtSmallerShape - make "Charlie" smaller
        elif idMso == 'SmartArtSmallerShape':
            _, sa = _find_smartart(doc)
            if sa:
                target = _find_node_by_text(sa, 'Charlie')
                if target:
                    target.Smaller()
                    target.Smaller()  # Apply twice for more noticeable change
                    return True
            return False

        # SmartArtRightToLeft - toggle reverse direction
        elif idMso == 'SmartArtRightToLeft':
            _, sa = _find_smartart(doc)
            if sa:
                # Toggle: if currently 0 (LTR), set -1 (RTL); if -1, set 0
                sa.Reverse = 0 if sa.Reverse else -1
                return True
            return False

        # SmartArtResetGraphic - reset SmartArt formatting
        elif idMso == 'SmartArtResetGraphic':
            _, sa = _find_smartart(doc)
            if sa:
                sa.Reset()
                return True
            return False

        return False

    except Exception as e:
        print(f"  API execution error for {idMso}: {e}")
        return False


# =============================================================================
# GENERATOR CLASS
# =============================================================================

class ObjectEffectContextualActionableGenerator:
    """
    Generator for object_effect × contextual_actionable test cases.
    
    This test set includes commands that operate on contextual objects:
    - Table commands (merge, split, alignment, etc.)
    - Picture commands (crop, reset, background removal, etc.)
    - SmartArt commands (add shapes, promote/demote, etc.)
    - Chart commands (format, refresh data, etc.)
    - WordArt/TextBox commands (alignment, 3D effects, etc.)
    - Equation commands (format conversion, etc.)
    - And more...
    
    All commands are:
    - BEHAVIOR: object_effect (verify via object property changes in XML)
    - SCOPE: contextual_actionable (setup: create object, execute: direct)
    
    Test strategy:
    - toggleButton: 2 tests (on, off)
    - button: 1 test (apply)
    
    Key characteristics:
    - Requires specific object types (tables, pictures, SmartArt, etc.)
    - Each command needs appropriate object setup
    - Selection varies by object type and command
    """

    def __init__(self, output_dir: Path = None):
        self.execution = DirectExecute()
        self.output_dir = output_dir or (BENCHMARK_OUTPUT_DIR / 'object_effect_x_contextual_actionable')
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.inputs_dir = self.output_dir / 'inputs'
        self.inputs_dir.mkdir(parents=True, exist_ok=True)

    def _needs_special_input(self, idMso: str) -> bool:
        """Check if command needs a special input document (not from shared inputs folder)."""
        special_commands = {
            # Table commands need specific table configurations
            'TableAutoFitContents',  # needs auto-fit to window
            'TableAutoFitWindow',  # needs narrow table
            'TableCellAlignMiddleCenter',  # needs top-left alignment
            'TableColumnsDistribute',  # needs uneven widths
            'TableRowsDistribute',  # needs uneven heights
            'TableStyleClear',  # needs table with style
            # Insert commands need narrower table so native behavior is deterministic
            'TableColumnsInsertLeft',  # needs narrow table (not page-width)
            'TableColumnsInsertRight',  # needs narrow table (not page-width)
            'TableRowsInsertAboveWord',  # needs narrow table (not page-width)
            'TableRowsInsertBelowWord',  # needs narrow table (not page-width)
            # Picture commands need modified pictures to show visual difference
            'PictureReset',  # needs picture with effects applied
            'PictureResetAndSize',  # needs picture with effects and size changes
            'PictureFillCrop',  # needs picture cropped differently
            'PictureFitCrop',  # needs picture cropped differently
            # Picture crop commands need NON-target aspect ratio input
            'PictureCropAspectRatio16To9',  # default is 16:9, needs different aspect ratio input
            'PictureCropAspectRatio4To3',  # needs wider image (16:9) to show crop effect
        }
        return idMso in special_commands

    def _generate_special_input_to_path(self, idMso: str, output_path: Path) -> Path:
        """
        Generate special input document directly to a specific path.
        Used for generating gold.docx in off tests for special input commands.

        Args:
            idMso: Command identifier
            output_path: Where to save the generated document

        Returns:
            Path to the generated document
        """
        # Detect variation from path (e.g., {command}/on/doc_in.docx -> 'on')
        variation = output_path.parent.name if output_path.parent.name in ['on', 'off', 'apply'] else None

        wd = ensure_word()
        try:
            if wd.Documents.Count > 0:
                for i in range(wd.Documents.Count, 0, -1):
                    try:
                        wd.Documents(i).Close(False)
                    except:
                        pass

            time.sleep(0.2)
            doc = wd.Documents.Add()
            time.sleep(0.2)

            # =============================================
            # Picture special inputs
            # =============================================

            # PictureReset/ResetAndSize - need picture with effects applied
            if idMso in ['PictureReset', 'PictureResetAndSize']:
                img_path = Path(tempfile.gettempdir()) / 'test_image_for_reset.png'
                try:
                    from PIL import ImageDraw
                    width, height = 640, 360
                    img = Image.new('RGB', (width, height), color=(73, 109, 137))
                    draw = ImageDraw.Draw(img)
                    corner_size = 60
                    draw.rectangle([0, 0, corner_size, corner_size], fill=(255, 80, 80))
                    draw.rectangle([width-corner_size, 0, width, corner_size], fill=(80, 200, 80))
                    draw.rectangle([0, height-corner_size, corner_size, height], fill=(80, 80, 255))
                    draw.rectangle([width-corner_size, height-corner_size, width, height], fill=(255, 220, 80))
                    center_x, center_y = width // 2, height // 2
                    draw.line([(center_x - 30, center_y), (center_x + 30, center_y)], fill=(255, 255, 255), width=3)
                    draw.line([(center_x, center_y - 30), (center_x, center_y + 30)], fill=(255, 255, 255), width=3)
                    draw.rectangle([0, 0, width-1, height-1], outline=(255, 255, 255), width=2)
                    img.save(str(img_path))

                    picture = doc.InlineShapes.AddPicture(
                        FileName=str(img_path), LinkToFile=False,
                        SaveWithDocument=True, Range=wd.Selection.Range
                    )
                    time.sleep(0.2)

                    # Apply VISIBLE formatting effects that PictureReset will clear
                    # PictureReset clears: brightness, contrast, color corrections, artistic effects
                    picture.Select()
                    time.sleep(0.1)

                    # Use API to set brightness and contrast (more reliable than SendKeys)
                    try:
                        picture.PictureFormat.Brightness = 0.3  # Make brighter
                        picture.PictureFormat.Contrast = 0.3    # Increase contrast
                        time.sleep(0.2)
                    except:
                        pass

                    # For ResetAndSize, also change the size
                    if idMso == 'PictureResetAndSize':
                        try:
                            picture.Width = picture.Width * 0.6
                            picture.Height = picture.Height * 0.6
                        except:
                            pass
                except Exception as e:
                    if 'rejected by callee' in str(e).lower() or '-2147418111' in str(e):
                        raise  # Re-raise COM busy errors
                    print(f"  Warning: Failed to create picture with effects: {e}")

            # PictureFillCrop/FitCrop - need picture cropped with content offset
            elif idMso in ['PictureFillCrop', 'PictureFitCrop']:
                img_path = Path(tempfile.gettempdir()) / 'test_image_for_fillfit.png'
                try:
                    from PIL import ImageDraw
                    # Create 16:9 image with markers
                    width, height = 640, 360
                    img = Image.new('RGB', (width, height), color=(73, 109, 137))
                    draw = ImageDraw.Draw(img)
                    corner_size = 60
                    draw.rectangle([0, 0, corner_size, corner_size], fill=(255, 80, 80))
                    draw.rectangle([width-corner_size, 0, width, corner_size], fill=(80, 200, 80))
                    draw.rectangle([0, height-corner_size, corner_size, height], fill=(80, 80, 255))
                    draw.rectangle([width-corner_size, height-corner_size, width, height], fill=(255, 220, 80))
                    center_x, center_y = width // 2, height // 2
                    draw.line([(center_x - 30, center_y), (center_x + 30, center_y)], fill=(255, 255, 255), width=3)
                    draw.line([(center_x, center_y - 30), (center_x, center_y + 30)], fill=(255, 255, 255), width=3)
                    draw.rectangle([0, 0, width-1, height-1], outline=(255, 255, 255), width=2)
                    img.save(str(img_path))

                    picture = doc.InlineShapes.AddPicture(
                        FileName=str(img_path), LinkToFile=False,
                        SaveWithDocument=True, Range=wd.Selection.Range
                    )
                    time.sleep(0.2)

                    # For Fill/Fit to have visible effect, crop with asymmetric offset
                    picture.Select()
                    time.sleep(0.1)
                    try:
                        picture.PictureFormat.CropLeft = 50  # Crop 50 points from left
                        picture.PictureFormat.CropRight = 100  # Crop 100 points from right
                        time.sleep(0.2)
                    except:
                        pass
                except Exception as e:
                    if 'rejected by callee' in str(e).lower() or '-2147418111' in str(e):
                        raise  # Re-raise COM busy errors
                    print(f"  Warning: Failed to create cropped picture: {e}")

            # PictureCropAspectRatio16To9 - need non-16:9 image
            elif idMso == 'PictureCropAspectRatio16To9':
                img_path = Path(tempfile.gettempdir()) / 'test_image_16x11.png'
                try:
                    from PIL import ImageDraw
                    width, height = 640, 440  # 16:11 aspect ratio
                    img = Image.new('RGB', (width, height), color=(73, 109, 137))
                    draw = ImageDraw.Draw(img)
                    corner_size = 60
                    draw.rectangle([0, 0, corner_size, corner_size], fill=(255, 80, 80))
                    draw.rectangle([width-corner_size, 0, width, corner_size], fill=(80, 200, 80))
                    draw.rectangle([0, height-corner_size, corner_size, height], fill=(80, 80, 255))
                    draw.rectangle([width-corner_size, height-corner_size, width, height], fill=(255, 220, 80))
                    center_x, center_y = width // 2, height // 2
                    draw.line([(center_x - 30, center_y), (center_x + 30, center_y)], fill=(255, 255, 255), width=3)
                    draw.line([(center_x, center_y - 30), (center_x, center_y + 30)], fill=(255, 255, 255), width=3)
                    draw.rectangle([0, 0, width-1, height-1], outline=(255, 255, 255), width=2)
                    img.save(str(img_path))

                    picture = doc.InlineShapes.AddPicture(
                        FileName=str(img_path), LinkToFile=False,
                        SaveWithDocument=True, Range=wd.Selection.Range
                    )
                    time.sleep(0.2)
                except Exception as e:
                    if 'rejected by callee' in str(e).lower() or '-2147418111' in str(e):
                        raise  # Re-raise COM busy errors
                    print(f"  Warning: Failed to create 16:11 picture: {e}")

            # PictureCropAspectRatio4To3 - need wider image (16:9) to show crop effect
            elif idMso == 'PictureCropAspectRatio4To3':
                img_path = Path(tempfile.gettempdir()) / 'test_image_16x9.png'
                try:
                    from PIL import ImageDraw
                    width, height = 640, 360  # 16:9 aspect ratio (wider than 4:3)
                    img = Image.new('RGB', (width, height), color=(73, 109, 137))
                    draw = ImageDraw.Draw(img)
                    corner_size = 60
                    draw.rectangle([0, 0, corner_size, corner_size], fill=(255, 80, 80))
                    draw.rectangle([width-corner_size, 0, width, corner_size], fill=(80, 200, 80))
                    draw.rectangle([0, height-corner_size, corner_size, height], fill=(80, 80, 255))
                    draw.rectangle([width-corner_size, height-corner_size, width, height], fill=(255, 220, 80))
                    center_x, center_y = width // 2, height // 2
                    draw.line([(center_x - 30, center_y), (center_x + 30, center_y)], fill=(255, 255, 255), width=3)
                    draw.line([(center_x, center_y - 30), (center_x, center_y + 30)], fill=(255, 255, 255), width=3)
                    draw.rectangle([0, 0, width-1, height-1], outline=(255, 255, 255), width=2)
                    img.save(str(img_path))

                    picture = doc.InlineShapes.AddPicture(
                        FileName=str(img_path), LinkToFile=False,
                        SaveWithDocument=True, Range=wd.Selection.Range
                    )
                    time.sleep(0.2)
                except Exception as e:
                    if 'rejected by callee' in str(e).lower() or '-2147418111' in str(e):
                        raise  # Re-raise COM busy errors
                    print(f"  Warning: Failed to create 16:9 picture: {e}")

            # =============================================
            # Table special inputs
            # =============================================
            else:
                # Create table (table special input commands)
                table = doc.Tables.Add(
                    Range=wd.Selection.Range,
                    NumRows=3,
                    NumColumns=3
                )
                for row in range(1, 4):
                    for col in range(1, 4):
                        table.Cell(row, col).Range.Text = f"R{row}C{col}"

                # Apply command-specific table settings
                if idMso == 'TableAutoFitContents':
                    # Table with auto-fit to window
                    try:
                        table.AutoFitBehavior(0)  # wdAutoFitFixed first
                        for col in table.Columns:
                            col.Width = 144
                        for row in table.Rows:
                            row.Height = 72
                            row.HeightRule = 1
                        table.AutoFitBehavior(2)  # wdAutoFitWindow
                    except:
                        pass

                elif idMso == 'TableAutoFitWindow':
                    # Narrow table with fixed widths
                    try:
                        table.AutoFitBehavior(0)  # wdAutoFitFixed
                        for col in table.Columns:
                            col.Width = 36  # Very narrow
                        for row in table.Rows:
                            row.Height = 72
                            row.HeightRule = 1
                    except:
                        pass

                elif idMso == 'TableCellAlignMiddleCenter':
                    # Table with top-left alignment (opposite of middle-center)
                    try:
                        table.AutoFitBehavior(0)  # wdAutoFitFixed
                        for col in table.Columns:
                            col.Width = 144
                        for row in table.Rows:
                            row.Height = 72
                            row.HeightRule = 1
                    except:
                        pass
                    # Keep default top-left alignment

                elif idMso in ['TableColumnsDistribute', 'TableRowsDistribute']:
                    # Table with uneven sizes
                    try:
                        table.AutoFitBehavior(0)  # wdAutoFitFixed
                        table.Columns(1).Width = 50
                        table.Columns(2).Width = 144
                        table.Columns(3).Width = 230
                        table.Rows(1).Height = 36
                        table.Rows(1).HeightRule = 1
                        table.Rows(2).Height = 72
                        table.Rows(2).HeightRule = 1
                        table.Rows(3).Height = 108
                        table.Rows(3).HeightRule = 1
                    except:
                        pass
                    # Set middle-center alignment
                    try:
                        for row in range(1, 4):
                            for col in range(1, 4):
                                cell = table.Cell(row, col)
                                cell.VerticalAlignment = 1
                                cell.Range.ParagraphFormat.Alignment = 1
                    except:
                        pass

                elif idMso == 'TableStyleClear':
                    # Table with style applied
                    try:
                        table.Style = "Table Grid"
                    except:
                        try:
                            table.Style = "Grid Table 1 Light"
                        except:
                            pass
                    try:
                        table.AutoFitBehavior(1)
                        for col in table.Columns:
                            col.Width = 144
                        for row in table.Rows:
                            row.Height = 72
                            row.HeightRule = 1
                    except:
                        pass
                    # Set middle-center alignment
                    try:
                        for row in range(1, 4):
                            for col in range(1, 4):
                                cell = table.Cell(row, col)
                                cell.VerticalAlignment = 1
                                cell.Range.ParagraphFormat.Alignment = 1
                    except:
                        pass

                # Add visible borders to make table structure clear for visual validation
                apply_table_borders(table)

            time.sleep(0.3)
            wd.Selection.HomeKey(Unit=6)

            doc.SaveAs(str(output_path.absolute()), FileFormat=16)
            time.sleep(0.1)
            doc.Close(False)
            time.sleep(0.2)

            return output_path

        except Exception as e:
            print(f"  Error creating special input: {e}")
            raise

    def generate_input_for_command(self, idMso: str, variation_name: Optional[str] = None) -> Path:
        """
        Generate input document with appropriate object for a command.

        Args:
            idMso: Command identifier
            variation_name: Variation name (apply, on, off) - used for special input path

        Returns:
            Path to the generated input document
        """
        precondition_type = get_precondition_type(idMso)
        input_filename = f'doc_in_{precondition_type}.docx'
        input_path = self.inputs_dir / input_filename

        # Special cases: input goes directly in command folder as doc_in.docx
        # (not in shared inputs folder)
        if self._needs_special_input(idMso) and variation_name:
            test_dir = self.output_dir / idMso / variation_name
            test_dir.mkdir(parents=True, exist_ok=True)
            input_path = test_dir / 'doc_in.docx'

            # Use _generate_special_input_to_path to create the special input
            if not input_path.exists():
                print(f"  Generating special input: {input_path.name}")
                return self._generate_special_input_to_path(idMso, input_path)
            return input_path

        if input_path.exists():
            return input_path

        print(f"  Generating input: {input_filename}")

        wd = ensure_word()
        try:
            # Ensure Word is in a clean state
            if wd.Documents.Count > 0:
                for i in range(wd.Documents.Count, 0, -1):
                    try:
                        wd.Documents(i).Close(False)
                    except:
                        pass

            time.sleep(0.2)
            doc = wd.Documents.Add()
            time.sleep(0.2)

            # Setup precondition (create appropriate object)
            setup_precondition(doc, wd, precondition_type, idMso)

            # Give Word time to process the changes
            time.sleep(0.3)

            # Ensure we're in a clean state before saving
            wd.Selection.HomeKey(Unit=6)  # wdStory - move to document start

            doc.SaveAs(str(input_path.absolute()), FileFormat=16)
            time.sleep(0.1)
            doc.Close(False)
            time.sleep(0.2)

            return input_path

        except Exception as e:
            print(f"  Error creating input: {e}")
            raise

    def load_commands(self) -> List[Dict]:
        """Load commands classified as object_effect × contextual_actionable."""
        commands_file = COMMAND_SETS_DIR / 'commands_by_policy_id.json'

        if not commands_file.exists():
            print(f"Warning: {commands_file} not found. Using empty command list.")
            return []

        with open(commands_file, 'r', encoding='utf-8') as f:
            all_commands = json.load(f)

        # Filter for our test set
        filtered = [
            cmd for cmd in all_commands
            if cmd.get('behavior') == 'object_effect'
            and cmd.get('scope') == 'contextual_actionable'
        ]

        # Ensure ALL_EXPANDED_COMMANDS are included (they may not be in the JSON file yet)
        existing_ids = {cmd['idMso'] for cmd in filtered}
        for idMso in ALL_EXPANDED_COMMANDS:
            if idMso not in existing_ids:
                filtered.append({
                    'idMso': idMso,
                    'behavior': 'object_effect',
                    'scope': 'contextual_actionable',
                    'control_type': 'button',
                })

        return filtered

    def get_variations_for_command(self, command: Dict) -> List[Dict]:
        """
        Get appropriate test variations based on control type.

        Args:
            command: Command dict with 'control_type' field

        Returns:
            List of variation dicts
        """
        control_type = command.get('control_type', 'button')

        if control_type == 'toggleButton':
            return TOGGLE_VARIATIONS
        else:
            return ONE_WAY_VARIATIONS

    def generate_on_test(self, idMso: str, variation: Dict, input_path: Path) -> Dict:
        """
        Generate 'on' test: apply effect to object.

        For toggle commands that set a specific state (e.g., EquationProfessionalAll),
        the input needs to be in the OPPOSITE state for the command to have an effect.
        
        Strategy:
        - doc_in: Start with OPPOSITE state (e.g., Linear for EquationProfessionalAll)
        - Execute the command to toggle TO the target state
        - gold: Result after applying command (e.g., Professional format)
        """
        test_dir = self.output_dir / idMso / variation['name']
        test_dir.mkdir(parents=True, exist_ok=True)

        gold_path = test_dir / 'gold.docx'
        spec_path = test_dir / 'spec.json'

        result = {
            'idMso': idMso,
            'variation': variation['name'],
            'success': False,
            'message': '',
            'paths': {
                'doc_in': str(input_path),
                'gold': str(gold_path),
                'spec': str(spec_path),
            }
        }

        # Skip check is now handled in generate_for_command()
        # No need to check again here

        wd = ensure_word()
        try:
            # Open input, apply command, save as gold
            doc = wd.Documents.Open(str(input_path.absolute()))
            time.sleep(0.3)

            # Select the appropriate object(s)
            precondition_type = get_precondition_type(idMso)
            if not select_for_command(doc, wd, self.execution, precondition_type, "TARGET", idMso):
                result['message'] = f"Failed to select object(s) for: {idMso}"
                doc.Close(False)
                return result

            # Execute the command - try API execution first for specific commands
            api_commands = {
                'HeaderFooterLinkToPrevious',
                'TableColumnsDistribute', 'TableRowsDistribute',
                'SplitCells',
                'TextDirection',
                'PictureCropAspectRatio1To1', 'PictureCropAspectRatio2To3', 'PictureCropAspectRatio3To2',
                'PictureCropAspectRatio3To4', 'PictureCropAspectRatio4To3', 'PictureCropAspectRatio3To5',
                'PictureCropAspectRatio5To3', 'PictureCropAspectRatio4To5', 'PictureCropAspectRatio5To4',
                'PictureCropAspectRatio16To9', 'PictureCropAspectRatio16To10',
                'PictureReset', 'PictureResetAndSize',
                'AlternativeText',
                # NOT PictureFitCrop/PictureFillCrop — they use ExecuteMso directly
                # Chart commands - ExecuteMso requires chart edit mode; use COM API instead
                'ChartSwitchRowColumn', 'ChartResetToMatchStyle',
                # SmartArt commands - node.Select() unavailable; use SmartArt node API
            } | SMARTART_COMMANDS

            if idMso in api_commands:
                if not execute_via_api(wd, doc, idMso):
                    result['message'] = f'Failed to execute command via API: {idMso}'
                    doc.Close(False)
                    return result
            else:
                # All other commands use ExecuteMso directly
                if not self.execution.execute(wd, idMso):
                    result['message'] = f'Failed to execute command: {idMso}'
                    doc.Close(False)
                    return result

            # NOTE: No post-command adjustments - test native Word command behavior
            # For insert commands, we use a narrow table (not page-width) so native insert
            # simply adds width/height without requiring redistribution

            # Save as gold
            doc.SaveAs(str(gold_path.absolute()), FileFormat=16)
            doc.Close(False)
            time.sleep(0.2)


            # Generate spec.json
            spec = self._generate_spec(idMso, variation, input_path, is_off_test=False)
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
        Generate 'off' test: document goes from effect-ON to effect-OFF.

        - doc_in: document WITH the effect applied (the "before" state)
        - gold: document AFTER the effect is removed/toggled off (the "after" state)

        Algorithm:
          1. Ensure doc_in exists with the effect applied
          2. Open doc_in, apply the "off" action, save as gold
        """
        test_dir = self.output_dir / idMso / variation['name']
        test_dir.mkdir(parents=True, exist_ok=True)

        doc_in_path = test_dir / 'doc_in.docx'
        gold_path = test_dir / 'gold.docx'
        spec_path = test_dir / 'spec.json'

        result = {
            'idMso': idMso,
            'variation': variation['name'],
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
            precondition_type = get_precondition_type(idMso)

            # =============================================
            # STEP 1: Create doc_in (with effect applied)
            # Open input, apply command, save as doc_in
            # =============================================
            doc = wd.Documents.Open(str(input_path.absolute()))
            time.sleep(0.3)

            if not select_for_command(doc, wd, self.execution, precondition_type, "TARGET", idMso):
                result['message'] = f"Failed to select object for doc_in: {idMso}"
                doc.Close(False)
                return result

            # Apply the effect to create the "ON" state
            api_commands = {
                'HeaderFooterLinkToPrevious',
                'TableColumnsDistribute', 'TableRowsDistribute',
            }
            if idMso in api_commands:
                if not execute_via_api(wd, doc, idMso):
                    result['message'] = f'Failed to apply effect for doc_in: {idMso}'
                    doc.Close(False)
                    return result
            else:
                if not self.execution.execute(wd, idMso):
                    result['message'] = f'Failed to apply effect for doc_in: {idMso}'
                    doc.Close(False)
                    return result

            time.sleep(0.2)
            doc.SaveAs(str(doc_in_path.absolute()), FileFormat=16)
            doc.Close(False)
            time.sleep(0.2)

            # =============================================
            # STEP 2: Create gold (effect OFF)
            # Three strategies based on command type:
            # =============================================

            # Strategy A: State-setter commands (equations) — execute opposite command
            off_state_commands = {
                'EquationProfessionalAll': 'EquationLinearAll',
                'EquationProfessionalOne': 'EquationLinearOne',
                'EquationLinearAll': 'EquationProfessionalAll',
                'EquationLinearOne': 'EquationProfessionalOne',
            }
            if idMso in off_state_commands:
                doc = wd.Documents.Open(str(doc_in_path.absolute()))
                time.sleep(0.3)
                if not select_for_command(doc, wd, self.execution, precondition_type, "TARGET", idMso):
                    result['message'] = f"Failed to select for gold: {idMso}"
                    doc.Close(False)
                    return result
                off_cmd = off_state_commands[idMso]
                if not execute_via_api(wd, doc, off_cmd):
                    if not self.execution.execute(wd, off_cmd):
                        result['message'] = f'Failed to execute off-state command: {off_cmd}'
                        doc.Close(False)
                        return result
                time.sleep(0.2)
                doc.SaveAs(str(gold_path.absolute()), FileFormat=16)
                doc.Close(False)
                time.sleep(0.2)

            # Strategy B: Special-input commands — regenerate original input as gold
            elif self._needs_special_input(idMso):
                self._generate_special_input_to_path(idMso, gold_path)

            # Strategy C: All other toggles (table alignment, SmartArtRightToLeft) —
            # gold = copy of shared input (precondition state)
            else:
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

    def generate_apply_test(self, idMso: str, variation: Dict, input_path: Path) -> Dict:
        """
        Generate 'apply' test for one-way (button) commands.
        Same logic as 'on' test.
        """
        return self.generate_on_test(idMso, variation, input_path)

    def _generate_spec(
        self,
        idMso: str,
        variation: Dict,
        input_path: Path,
        is_off_test: bool = False
    ) -> Dict:
        """Generate the spec.json for a test case."""

        variation_name = variation['name']
        precondition_type = get_precondition_type(idMso)

        # Calculate relative path to input
        if is_off_test:
            doc_in_ref = 'doc_in.docx'  # Local file for 'off' test
        elif self._needs_special_input(idMso):
            doc_in_ref = 'doc_in.docx'  # Special input is in command folder
        else:
            # Use the shared input from inputs folder
            input_filename = input_path.name
            doc_in_ref = f'../../inputs/{input_filename}'

        spec = {
            'idMso': idMso,
            'behavior': 'object_effect',
            'scope': 'contextual_actionable',
            'variation': variation_name,
            'generated_at': datetime.now().isoformat(),

            'precondition': {
                'type': precondition_type,
                'state': variation.get('precondition', 'base'),
            },

            'verification': {
                'primary': {
                    'method': 'shape_property_comparison',
                    'description': 'Word COM shape/object property comparison',
                },
                'invariants': [
                    {
                        'method': 'structure_preserved' if idMso in TEXT_CHANGING_COMMANDS else 'text_preserved',
                        'description': (
                            'Document structure preserved (text content may change)'
                            if idMso in TEXT_CHANGING_COMMANDS
                            else 'Text content preserved (object properties may change)'
                        ),
                    }
                ],
            },

            'files': {
                'doc_in': doc_in_ref,
                'gold': 'gold.docx',
            }
        }

        return spec

    # Commands to skip during generation (code kept but no test cases produced)
    SKIP_COMMANDS = {'AlternativeText'}

    def generate_for_command(self, command: Dict) -> List[Dict]:
        """
        Generate all test variations for a single command.

        Args:
            command: Command dict with 'idMso' and 'control_type'

        Returns:
            List of generation results
        """
        idMso = command['idMso']
        if idMso in self.SKIP_COMMANDS:
            print(f"  Skipping {idMso} (in SKIP_COMMANDS)")
            return []
        variations = self.get_variations_for_command(command)

        # Check if command should be skipped (before generating input)
        if idMso in SPECIAL_HANDLING_COMMANDS:
            skip_reason = SPECIAL_HANDLING_COMMANDS[idMso]
            results = []
            for var in variations:
                test_dir = self.output_dir / idMso / var['name']
                test_dir.mkdir(parents=True, exist_ok=True)
                spec_path = test_dir / 'spec.json'
                
                result = {
                    'idMso': idMso,
                    'variation': var['name'],
                    'success': False,
                    'skipped': True,
                    'message': f'Skipped: {skip_reason}',
                    'paths': {'spec': str(spec_path)},
                }
                
                # Generate spec documenting the skip
                spec = self._generate_spec(idMso, var, Path(''), is_off_test=False)
                spec['skipped'] = True
                spec['skip_reason'] = skip_reason
                write_spec_if_changed(spec_path, spec)
                
                results.append(result)
            
            return results

        # Generate input document for this command
        # For special input commands, input is generated per-variation in the test folder
        # For regular commands, input is shared in the inputs folder
        if not self._needs_special_input(idMso):
            input_path = self.generate_input_for_command(idMso)
        else:
            input_path = None  # Will be generated per-variation

        results = []
        for var in variations:
            # For special input commands, generate input per variation
            if self._needs_special_input(idMso):
                input_path = self.generate_input_for_command(idMso, var['name'])
            
            if var['name'] == 'on':
                result = self.generate_on_test(idMso, var, input_path)
            elif var['name'] == 'off':
                result = self.generate_off_test(idMso, var, input_path)
            elif var['name'] == 'apply':
                result = self.generate_apply_test(idMso, var, input_path)
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
        print(f"Test set: object_effect × contextual_actionable")
        print(f"Output: {self.output_dir}\n")

        summary = {
            'test_set': 'object_effect_x_contextual_actionable',
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

        # Print detailed results
        print(f"\n=== Generation Complete ===")
        print(f"\nResults\n")
        print("| Status     | Count |")
        print("|-----------|-------|")
        print(f"| Successful| {summary['successful']:<5} |")
        print(f"| Skipped   | {summary.get('skipped', 0):<5} |")
        print(f"| Failed    | {summary['failed']:<5} |")

        # Print generated test cases summary
        successful_results = [r for r in summary['results'] if r['success']]
        skipped_results = [r for r in summary['results'] if r.get('skipped')]

        if successful_results or skipped_results:
            print(f"\nGenerated Test Cases ({len(successful_results)} successful, {len(skipped_results)} skipped)\n")
            all_generated = successful_results + skipped_results
            for r in all_generated[:10]:  # Show first 10
                status = 'SKIP' if r.get('skipped') else 'OK'
                print(f"  [{status}] {r['idMso']}/{r['variation']}")
            if len(all_generated) > 10:
                print(f"  ... and {len(all_generated) - 10} more")

        print(f"\nTotal commands: {summary['total_commands']}")
        print(f"  - toggleButton: {summary['toggle_commands']} (2 tests each)")
        print(f"  - button: {summary['button_commands']} (1 test each)")
        print(f"Total test cases: {summary['total_test_cases']}")
        print(f"Summary saved to: {summary_path}")

        return summary


def main():
    """CLI entry point for generating object_effect × contextual_actionable tests."""
    import argparse

    parser = argparse.ArgumentParser(
        description='Generate test cases for object_effect × contextual_actionable'
    )
    parser.add_argument(
        '--command', '-c',
        help='Generate for a specific command (e.g., MergeCells)'
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
    generator = ObjectEffectContextualActionableGenerator(output_dir=output_dir)

    if args.command:
        # Generate for a single command
        commands = generator.load_commands()
        cmd = next((c for c in commands if c['idMso'] == args.command), None)

        if cmd:
            print(f"Generating test cases for {args.command} ({cmd.get('control_type', 'button')})...")
            results = generator.generate_for_command(cmd)
            success = sum(1 for r in results if r['success'])
            print(f"\nDone: {success}/{len(results)} successful")
        else:
            print(f"Command {args.command} not found in object_effect × contextual_actionable set.")
            print(f"Generating as button (1 test)...")
            cmd = {'idMso': args.command, 'control_type': 'button'}
            results = generator.generate_for_command(cmd)
            success = sum(1 for r in results if r['success'])
            print(f"\nDone: {success}/{len(results)} successful")
    else:
        # Generate for all commands
        generator.generate_all(limit=args.limit)


if __name__ == '__main__':
    main()
