"""Configuration constants for benchmark generation."""

from pathlib import Path

# Paths
PROJECT_ROOT = Path(__file__).parent.parent
COMMAND_SETS_DIR = PROJECT_ROOT / 'command_classification'
COMMAND_SPECS_DIR = PROJECT_ROOT / 'command_specs'
BENCHMARK_OUTPUT_DIR = PROJECT_ROOT / 'benchmark'

# XML Namespaces for OpenXML
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    'wpg': 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
}

# Word COM constants
WD_STORY = 6  # wdStory - entire document
WD_COLLAPSE_END = 0  # wdCollapseEnd

# Test variations based on control type
TOGGLE_VARIATIONS = [
    {
        'name': 'on',
        'description': 'Turn formatting on',
        'precondition': 'unformatted',
    },
    {
        'name': 'off',
        'description': 'Turn formatting off',
        'precondition': 'formatted',
    },
]

ONE_WAY_VARIATIONS = [
    {
        'name': 'apply',
        'description': 'Apply the action',
        'precondition': 'unformatted',
    },
]

# Fixed fill values for content control fields in gallery building blocks
# (Header, Footer, CoverPage). Used by both discovery scripts and generators.
FILL_VALUES = {
    'Title':         'Quarterly Business Review',
    'Subtitle':      'Strategic Planning Summary',
    'Author':        'Megan Bowen',
    'Company':       'Contoso Ltd.',
    'Date':          '2026-03-25',
    'Year':          '2026',
    'Abstract':      'A brief summary of the quarterly business review and strategic initiatives.',
    'Address':       '123 Main Street, Redmond, WA 98052',
    'Email':         'megan@contoso.com',
    'Publish Date':  '2026-03-25',
    'Course':        'Business Strategy 101',
    'School':        'Contoso University',
}

# Case-scoped, non-identifying source for DocInfoAuthor.  Keep this separate
# from FILL_VALUES['Author']: other fixtures still use the long-standing sample
# persona, while DocInfoAuthor must expose a solver-visible synthetic source
# that survives the export identity sanitizer.
DOCINFO_AUTHOR_FIXTURE = 'WordEval Author'
