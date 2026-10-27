"""Setup templates derived from SCOPE dimension."""

import os
import time
from typing import Optional

try:
    import win32com.client as win32
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False
    print("Warning: win32com not available. Word automation will not work.")


def ensure_word(visible: bool = True):
    """
    Get or create Word application instance.

    Args:
        visible: If True, show Word window. Default True because some operations
                 (e.g., IncrementRotation) require Word to be visible.
    """
    if not HAS_WIN32:
        raise RuntimeError("win32com is required for Word automation")

    try:
        wd = win32.GetActiveObject("Word.Application")
    except Exception:
        wd = win32.Dispatch("Word.Application")

    wd.Visible = visible
    wd.DisplayAlerts = 0  # wdAlertsNone - suppress dialogs
    return wd


def finalize_word(wd, doc=None):
    """Clean up Word resources."""
    if doc:
        try:
            doc.Close(False)  # Don't save changes
        except Exception:
            pass
    # Don't quit Word - may be reused


class SetupTemplate:
    """Base class for setup templates."""

    def __init__(self):
        self.wd = None
        self.doc = None

    def setup(self, **kwargs) -> str:
        """Create document and return path. Override in subclasses."""
        raise NotImplementedError

    def cleanup(self):
        """Clean up resources."""
        finalize_word(self.wd, self.doc)


class SimpleTextSetup(SetupTemplate):
    """
    Setup for core_actionable and core_parameterized scopes.

    Creates a simple text document with selectable target text.
    """

    def setup(self, full_text: str, output_path: str, **kwargs) -> str:
        """
        Create a simple text document.

        Args:
            full_text: The text content to insert
            output_path: Where to save the .docx file

        Returns:
            Absolute path to the created document
        """
        self.wd = ensure_word()
        self.doc = self.wd.Documents.Add()

        # Type the text
        self.wd.Selection.TypeText(full_text)

        # Save
        abs_path = os.path.abspath(output_path)
        self.doc.SaveAs(abs_path, FileFormat=16)  # 16 = wdFormatDocumentDefault (.docx)

        return abs_path


class MultiParagraphSetup(SetupTemplate):
    """
    Setup for paragraph formatting commands.

    Creates a document with multiple paragraphs.
    """

    def setup(self, paragraphs: list, output_path: str, **kwargs) -> str:
        """
        Create a multi-paragraph document.

        Args:
            paragraphs: List of paragraph texts
            output_path: Where to save the .docx file

        Returns:
            Absolute path to the created document
        """
        self.wd = ensure_word()
        self.doc = self.wd.Documents.Add()

        for i, para in enumerate(paragraphs):
            self.wd.Selection.TypeText(para)
            if i < len(paragraphs) - 1:
                self.wd.Selection.TypeParagraph()

        abs_path = os.path.abspath(output_path)
        self.doc.SaveAs(abs_path, FileFormat=16)

        return abs_path


class ObjectSetup(SetupTemplate):
    """
    Setup for contextual_actionable and contextual_parameterized scopes.

    Creates a document with an object (table, picture, shape, etc.).
    """

    def setup(self, object_type: str, output_path: str, **kwargs) -> str:
        """
        Create a document with an object.

        Args:
            object_type: Type of object ('table', 'picture', 'shape', etc.)
            output_path: Where to save the .docx file

        Returns:
            Absolute path to the created document
        """
        self.wd = ensure_word()
        self.doc = self.wd.Documents.Add()

        if object_type == 'table':
            # Insert a 3x3 table
            self.doc.Tables.Add(
                Range=self.wd.Selection.Range,
                NumRows=3,
                NumColumns=3
            )
        elif object_type == 'picture':
            # Insert a placeholder shape (for testing)
            image_path = kwargs.get('image_path', '')
            if image_path:
                self.wd.Selection.InlineShapes.AddPicture(
                    FileName=image_path,
                    LinkToFile=False,
                    SaveWithDocument=True
                )
        elif object_type == 'shape':
            # Insert a rectangle shape
            self.doc.Shapes.AddShape(
                Type=1,  # msoShapeRectangle
                Left=100,
                Top=100,
                Width=200,
                Height=100
            )

        abs_path = os.path.abspath(output_path)
        self.doc.SaveAs(abs_path, FileFormat=16)

        return abs_path


# Factory function
def get_setup_template(scope: str) -> SetupTemplate:
    """Get the appropriate setup template for a SCOPE."""

    setup_map = {
        'core_actionable': SimpleTextSetup,
        'core_parameterized': SimpleTextSetup,
        'core_dialog': SimpleTextSetup,
        'specialized_actionable': SimpleTextSetup,
        'specialized_parameterized': SimpleTextSetup,
        'contextual_actionable': ObjectSetup,
        'contextual_parameterized': ObjectSetup,
        'contextual_dialog': ObjectSetup,
    }

    template_class = setup_map.get(scope, SimpleTextSetup)
    return template_class()
