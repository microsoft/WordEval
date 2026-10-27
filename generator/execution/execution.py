"""Execution templates derived from SCOPE dimension."""

import time
from typing import Optional


class ExecutionTemplate:
    """Base class for execution templates."""

    def select(self, wd, **kwargs) -> bool:
        """Apply selection. Override in subclasses."""
        raise NotImplementedError

    def execute(self, wd, idMso: str, **kwargs) -> bool:
        """Execute command. Override in subclasses."""
        raise NotImplementedError


class DirectExecute(ExecutionTemplate):
    """
    Execution for *_actionable scopes.

    Selects content then executes command directly via ExecuteMso.
    """

    def select(self, wd, select_text: str, **kwargs) -> bool:
        """
        Find and select specific text in the document.

        Args:
            wd: Word application instance
            select_text: Text to find and select

        Returns:
            True if text found and selected
        """
        # Go to document start
        wd.Selection.HomeKey(Unit=6)  # wdStory = 6

        # Find and select
        wd.Selection.Find.ClearFormatting()
        found = wd.Selection.Find.Execute(
            FindText=select_text,
            MatchCase=False,
            MatchWholeWord=False,
            Forward=True
        )

        return found

    def execute(self, wd, idMso: str, **kwargs) -> bool:
        """
        Execute command via ExecuteMso.

        Args:
            wd: Word application instance
            idMso: The command identifier

        Returns:
            True if execution succeeded
        """
        try:
            wd.CommandBars.ExecuteMso(idMso)
            time.sleep(0.1)  # Brief pause for command to complete
            return True
        except Exception as e:
            print(f"ExecuteMso failed for {idMso}: {e}")
            return False


class ParameterizedExecute(ExecutionTemplate):
    """
    Execution for *_parameterized scopes.

    Requires selecting a parameter value before execution.
    """

    def select(self, wd, select_text: str, **kwargs) -> bool:
        """Same as DirectExecute."""
        wd.Selection.HomeKey(Unit=6)
        wd.Selection.Find.ClearFormatting()
        return wd.Selection.Find.Execute(
            FindText=select_text,
            MatchCase=False,
            MatchWholeWord=False,
            Forward=True
        )

    def execute(self, wd, idMso: str, parameter_value: str = None, **kwargs) -> bool:
        """
        Execute parameterized command.

        For galleries/dropdowns, need to select the parameter value first.
        This is more complex and may require UI automation.
        """
        # TODO: Implement parameterized execution
        # This may involve:
        # 1. Opening the gallery/dropdown
        # 2. Selecting the specific item
        # 3. Or using specific ExecuteMso for predefined values
        raise NotImplementedError("Parameterized execution not yet implemented")


class DialogExecute(ExecutionTemplate):
    """
    Execution for *_dialog scopes.

    Opens a dialog, fills in values, and confirms.
    """

    def select(self, wd, select_text: str = None, **kwargs) -> bool:
        """Select content if needed."""
        if select_text:
            wd.Selection.HomeKey(Unit=6)
            wd.Selection.Find.ClearFormatting()
            return wd.Selection.Find.Execute(
                FindText=select_text,
                MatchCase=False,
                MatchWholeWord=False,
                Forward=True
            )
        return True

    def execute(self, wd, idMso: str, dialog_actions: list = None, **kwargs) -> bool:
        """
        Execute dialog command.

        Opens the dialog and performs specified actions.
        May require SendKeys or UI Automation.
        """
        # TODO: Implement dialog automation
        raise NotImplementedError("Dialog execution not yet implemented")


# Factory function
def get_execution_template(scope: str) -> ExecutionTemplate:
    """Get the appropriate execution template for a SCOPE."""

    if scope.endswith('_actionable'):
        return DirectExecute()
    elif scope.endswith('_parameterized'):
        return ParameterizedExecute()
    elif scope.endswith('_dialog'):
        return DialogExecute()
    else:
        return DirectExecute()
