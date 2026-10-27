"""Shared utilities for interaction-sensitive composition generators.

Provides object creation helpers, input document management, and spec
construction utilities used across all interaction categories.
"""

import json
import shutil
import time
from datetime import datetime
from pathlib import Path

from generator.config import BENCHMARK_OUTPUT_DIR


# ── Asset paths ──────────────────────────────────────────────────────────
ASSETS_DIR = Path(__file__).parent.parent.parent / 'assets'
IMG_PATH = str(ASSETS_DIR / 'img.png')


# ── Directories that contain reusable input documents ────────────────────
INPUT_SEARCH_DIRS = [
    BENCHMARK_OUTPUT_DIR / 'object_effect_x_contextual_parameterized' / 'inputs',
    BENCHMARK_OUTPUT_DIR / 'object_effect_x_core_parameterized' / 'inputs',
    BENCHMARK_OUTPUT_DIR / 'object_effect_x_core_actionable' / 'inputs',
    BENCHMARK_OUTPUT_DIR / 'object_effect_x_contextual_actionable' / 'inputs',
    BENCHMARK_OUTPUT_DIR / 'content_operations_x_contextual_actionable' / 'inputs',
    BENCHMARK_OUTPUT_DIR / 'content_formatting_x_core_parameterized' / 'inputs',
    BENCHMARK_OUTPUT_DIR / 'additive_composition_x_content_formatting' / 'inputs',
    BENCHMARK_OUTPUT_DIR / 'additive_composition_x_object_effect' / 'inputs',
]


# ═════════════════════════════════════════════════════════════════════════
# Object creation helpers
# ═════════════════════════════════════════════════════════════════════════

def create_table(doc, rows=3, cols=3, fill_cells=True):
    """Insert a table with optional cell labels (R1C1..RnCn).

    Returns the Table COM object.
    """
    rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    table = doc.Tables.Add(rng, rows, cols)
    time.sleep(0.3)
    if fill_cells:
        for r in range(1, rows + 1):
            for c in range(1, cols + 1):
                table.Cell(r, c).Range.Text = f'R{r}C{c}'
    time.sleep(0.2)
    return table


def _select_end_of_doc(doc):
    """Collapse the selection to the end of the document.

    Inline-shape insertion APIs (AddChart2, AddPicture, ExecuteMso for
    equation) insert at the current selection.  When a document opens,
    the selection defaults to position 0 — which in contextual stub docs
    would place the inserted object BEFORE any heading content.  This
    helper moves the insertion point to end-of-document so inserts
    appear after existing paragraphs.
    """
    end_range = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    end_range.Select()
    time.sleep(0.1)


def create_chart(doc, chart_type=51, setup_for_testing=True,
                 categories=None, series=None):
    """Insert an inline chart at the end of the document.

    Args:
        chart_type: Excel chart type constant (51 = xlColumnClustered,
                    4 = xlLine).
        setup_for_testing: If True, turn off title and gridlines so
                           composition steps can enable them and show a diff.
        categories: Optional list of category labels (e.g. ['Q1', 'Q2']).
                    If provided, chart data is set explicitly instead of
                    using Word's template defaults — making the gold
                    deterministic and tool-independent.
        series: Optional list of dicts with 'name' and 'values' keys.
                E.g. [{'name': 'Revenue', 'values': [100, 150]}].
                Must match the length of categories.

    Returns the Chart COM object.
    """
    _select_end_of_doc(doc)
    doc.InlineShapes.AddChart2(-1, chart_type)
    time.sleep(3)
    chart = doc.InlineShapes(1).Chart

    # Set explicit data if provided (replaces Word's template defaults).
    # Uses the SeriesCollection API directly — avoids activating the
    # embedded Excel workbook, which is fragile under COM automation.
    if categories and series:
        # Remove excess default series (Word default is 3 series)
        while chart.SeriesCollection().Count > len(series):
            chart.SeriesCollection(chart.SeriesCollection().Count).Delete()
            time.sleep(0.2)
        # Set data for each series
        for i, s in enumerate(series, 1):
            if i <= chart.SeriesCollection().Count:
                sc = chart.SeriesCollection(i)
            else:
                sc = chart.SeriesCollection().NewSeries()
            sc.Name = s['name']
            sc.Values = s['values']
            sc.XValues = categories
            time.sleep(0.2)

    if setup_for_testing:
        chart.HasTitle = False
        for ax_type in [1, 2]:  # xlCategory, xlValue
            try:
                ax = chart.Axes(ax_type)
                ax.HasMajorGridlines = False
                ax.HasMinorGridlines = False
            except Exception:
                pass
    time.sleep(0.5)
    return chart


_ICON_SVGS = {
    'star': (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 200"'
        ' width="200" height="200">'
        '<polygon points="100,10 127,80 200,80 140,125 160,195'
        ' 100,150 40,195 60,125 0,80 73,80"'
        ' fill="#4472C4" stroke="#2F5496" stroke-width="2"/>'
        '</svg>'
    ),
    'circle': (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 200"'
        ' width="200" height="200">'
        '<circle cx="100" cy="100" r="90"'
        ' fill="#ED7D31" stroke="#C55A11" stroke-width="2"/>'
        '</svg>'
    ),
}
ICON_SIZE_PT = 150.0


def create_icon(doc, shape='star'):
    """Insert an inline SVG icon at the end of the document.

    Generates a simple SVG graphic, writes to a temp file, inserts via
    InlineShapes.AddPicture, and deletes the temp file.  No external
    file dependency — the icon is self-contained.

    Args:
        shape: 'star' or 'circle'.

    Returns the InlineShape COM object.
    """
    import tempfile

    svg_content = _ICON_SVGS.get(shape, _ICON_SVGS['star'])
    tmp = tempfile.NamedTemporaryFile(
        suffix='.svg', delete=False, mode='w', encoding='utf-8'
    )
    tmp.write(svg_content)
    tmp.close()

    _select_end_of_doc(doc)
    pic = doc.InlineShapes.AddPicture(tmp.name)
    pic.Width = ICON_SIZE_PT
    pic.Height = ICON_SIZE_PT
    time.sleep(0.5)
    Path(tmp.name).unlink(missing_ok=True)
    return pic


def create_picture(doc, img_path=None):
    """Insert an inline picture from img.png at the end of the document.

    Returns the InlineShape COM object.
    """
    if img_path is None:
        img_path = IMG_PATH
    _select_end_of_doc(doc)
    pic = doc.InlineShapes.AddPicture(img_path)
    time.sleep(0.5)
    return pic


def create_landscape_picture(doc, width=640, height=480):
    """Insert a landscape picture (4:3) with colored quadrant markers.

    4:3 ratio ensures both 16:9 and 1:1 crops produce visible changes.
    Colored corners make crop direction verifiable:
    red=top-left, green=top-right, blue=bottom-left, yellow=bottom-right.

    Returns the InlineShape COM object.
    """
    import tempfile
    from PIL import Image, ImageDraw

    img = Image.new('RGB', (width, height), (100, 160, 220))
    draw = ImageDraw.Draw(img)
    q = min(width, height) // 4
    draw.rectangle([0, 0, q, q], fill=(220, 60, 60))           # red TL
    draw.rectangle([width - q, 0, width, q], fill=(60, 180, 60))  # green TR
    draw.rectangle([0, height - q, q, height], fill=(60, 60, 220))  # blue BL
    draw.rectangle([width - q, height - q, width, height],
                   fill=(220, 200, 60))                          # yellow BR
    # Center cross
    cx, cy = width // 2, height // 2
    draw.line([(cx - 20, cy), (cx + 20, cy)], fill='white', width=2)
    draw.line([(cx, cy - 20), (cx, cy + 20)], fill='white', width=2)

    tmp = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
    img.save(tmp.name)
    tmp.close()
    pic = doc.InlineShapes.AddPicture(tmp.name)
    time.sleep(0.5)
    Path(tmp.name).unlink(missing_ok=True)
    return pic


def create_shape(doc, shape_type=1, left=100, top=100, width=100, height=80):
    """Insert a floating shape.

    Args:
        shape_type: MsoAutoShapeType (1 = rectangle).

    Returns the Shape COM object.
    """
    shape = doc.Shapes.AddShape(shape_type, left, top, width, height)
    time.sleep(0.2)
    return shape


def get_smartart_layout(wd, *, layout_index=None, layout_name=None):
    """Resolve a SmartArt layout by stable display name when one is required."""
    layouts = wd.Application.SmartArtLayouts
    if layout_name:
        expected = str(layout_name).strip().casefold()
        for index in range(1, layouts.Count + 1):
            layout = layouts(index)
            if str(layout.Name or '').strip().casefold() == expected:
                return layout
        raise ValueError(f'SmartArt layout not found by name: {layout_name}')
    if layout_index is None:
        raise ValueError('layout_index or layout_name is required')
    return layouts(int(layout_index))


def create_smartart(wd, doc, layout_index=1, labels=None,
                    left=50, top=50, width=400, height=300,
                    layout_name=None):
    """Insert a SmartArt graphic.

    Args:
        layout_index: 1-based index into SmartArtLayouts
                      (1 = Basic Block List, 102 = Hierarchy).
        labels: List of strings for node text. If None, default labels
                are applied.

    Returns the Shape COM object containing the SmartArt.
    """
    layout = get_smartart_layout(
        wd, layout_index=layout_index, layout_name=layout_name
    )
    sa_shape = doc.Shapes.AddSmartArt(layout, left, top,
                                      width, height)
    time.sleep(1.5)
    sa = sa_shape.SmartArt
    if labels is not None:
        nodes = sa.AllNodes
        for i in range(1, min(nodes.Count + 1, len(labels) + 1)):
            nodes(i).TextFrame2.TextRange.Text = labels[i - 1]
    time.sleep(0.3)
    return sa_shape


def create_equation(wd, doc, content='y='):
    """Insert an equation at the end of the document and type content into it.

    Returns nothing -- the equation is selected after creation.
    """
    _select_end_of_doc(doc)
    wd.CommandBars.ExecuteMso('EquationInsertNew')
    time.sleep(1.0)
    wd.Selection.TypeText(content)
    time.sleep(0.3)


def create_textbox(doc, text='Sample Text',
                   left=100, top=100, width=200, height=80):
    """Insert a text box with specified text.

    Returns the Shape COM object.
    """
    tb = doc.Shapes.AddTextbox(1, left, top, width, height)  # msoTextOrientationHorizontal
    time.sleep(0.3)
    tb.TextFrame.TextRange.Text = text
    time.sleep(0.2)
    return tb


# ═════════════════════════════════════════════════════════════════════════
# Input document management
# ═════════════════════════════════════════════════════════════════════════

def get_shared_input(inputs_dir, input_name):
    """Copy a shared input document into the test set's inputs/ directory.

    Searches INPUT_SEARCH_DIRS for ``{input_name}.docx`` and copies it
    to ``inputs_dir/{input_name}.docx``.  Returns the destination path.

    Raises FileNotFoundError if the input is not found anywhere.
    """
    dest = inputs_dir / f'{input_name}.docx'
    if dest.exists():
        return dest

    for src_dir in INPUT_SEARCH_DIRS:
        src = src_dir / f'{input_name}.docx'
        if src.exists():
            shutil.copy2(src, dest)
            print(f'  Copied input: {input_name}.docx')
            return dest

    raise FileNotFoundError(
        f"Shared input '{input_name}.docx' not found in any benchmark "
        f"inputs directory"
    )


def create_empty_doc(wd, dest_path):
    """Create a near-empty document (single empty paragraph) and save it."""
    doc = wd.Documents.Add()
    time.sleep(0.3)
    doc.SaveAs(str(Path(dest_path).absolute()), FileFormat=16)
    doc.Close(False)
    time.sleep(0.3)


def create_stub_doc(wd, dest_path, heading_text: str):
    """Create a stub document that provides semantic context for an insertion.

    The document contains a single heading line describing the purpose of
    the object that an agent will insert (e.g., "Quarterly Sales:" for a
    chart, "Team Structure:" for a SmartArt diagram).  This replaces the
    sterile empty-document precondition with a contextual anchor — the
    heading names *why* the insertion is happening, so the agent can see
    a document-level intent rather than an empty canvas.

    Verification is unaffected: the inserted object is still the only
    shape/table/chart in the document, and verifiers compare object
    properties (not surrounding paragraph text).
    """
    doc = wd.Documents.Add()
    time.sleep(0.3)
    # Apply a Heading 2 style to the placeholder line so it visually
    # reads as a section header rather than body text.
    rng = doc.Content
    rng.Text = heading_text + '\r'
    try:
        rng.Paragraphs(1).Style = 'Heading 2'
    except Exception:
        pass
    # Move insertion point to the end so agents inserting at cursor
    # position naturally place the object below the heading.
    doc.Content.Collapse(0)  # wdCollapseEnd
    doc.SaveAs(str(Path(dest_path).absolute()), FileFormat=16)
    doc.Close(False)
    time.sleep(0.3)


def create_tall_shapes_doc(wd, dest_path):
    """Create a document with three tall narrow rectangles (60x120).

    Shapes are staggered diagonally at (50,50), (150,100), (250,150)
    so rotation visibly changes their bounding boxes.
    """
    doc = wd.Documents.Add()
    time.sleep(0.3)
    positions = [(50, 50), (150, 100), (250, 150)]
    for left, top in positions:
        doc.Shapes.AddShape(1, left, top, 60, 120)  # Rectangle
        time.sleep(0.2)
    doc.SaveAs(str(Path(dest_path).absolute()), FileFormat=16)
    doc.Close(False)
    time.sleep(0.3)


def create_spread_tall_shapes_doc(wd, dest_path):
    """Create a document with three tall narrow rectangles (60x120), spread unevenly.

    Shapes at (50,50), (170,200), (400,550) — UNEVEN in both axes so that
    distribute produces visible changes after rotation changes bounding box.
    Middle shape is closer to shape 1 than to shape 3 on both axes.
    """
    doc = wd.Documents.Add()
    time.sleep(0.3)
    positions = [(50, 50), (170, 200), (400, 550)]
    for left, top in positions:
        doc.Shapes.AddShape(1, left, top, 60, 120)  # Rectangle
        time.sleep(0.2)
    doc.SaveAs(str(Path(dest_path).absolute()), FileFormat=16)
    doc.Close(False)
    time.sleep(0.3)


# ═════════════════════════════════════════════════════════════════════════
# Spec construction
# ═════════════════════════════════════════════════════════════════════════

def build_spec(*, behavior, composition_type, interaction_class,
               target_object_type, steps, precondition,
               verification_method, verification_description,
               doc_in_ref, variant_name, auxiliary=None,
               invariants=None, atomic_intents=None):
    """Construct a spec dict for an interaction composition test case.

    Args:
        behavior: 'object_effect', 'content_formatting', or
                  'content_operations'.
        composition_type: 'prerequisite', 'scope_changing',
                          'layout_coupled', or 'overwrite'.
        interaction_class: Fine-grained tag, e.g. 'insert_then_style'.
        target_object_type: 'table', 'chart', 'shape', etc.
        steps: List of step dicts, each with 'operation' or 'idMso'
               and optional 'params'.
        precondition: Dict describing input state.
        verification_method: String for verification.primary.method.
        verification_description: Human description.
        doc_in_ref: Relative path to doc_in from the variant dir.
        variant_name: e.g. 'variant_1'.
        auxiliary: Optional list of auxiliary file paths.
        invariants: Optional list of invariant dicts.
        atomic_intents: Optional list of intent-level decomposition metadata.

    Returns:
        Dict ready for JSON serialization.
    """
    spec = {
        'behavior': behavior,
        'scope': 'interaction_composition',
        'variation': variant_name,
        'generated_at': datetime.now().isoformat(),
        'composition': {
            'type': composition_type,
            'steps': steps,
            'operation_count': len(steps),
            'interaction_class': interaction_class,
            'target_object_type': target_object_type,
        },
        'precondition': precondition,
        'verification': {
            'primary': {
                'method': verification_method,
                'description': verification_description,
            },
        },
        'files': {
            'doc_in': doc_in_ref,
            'gold': 'gold.docx',
        },
    }

    if auxiliary:
        spec['files']['auxiliary'] = auxiliary

    if invariants:
        spec['verification']['invariants'] = invariants

    if atomic_intents:
        spec['composition']['atomic_intents'] = atomic_intents

    return spec


def write_spec(spec_path, spec):
    """Write spec.json, skipping if content is unchanged (ignoring timestamps)."""
    from generator.test_sets.spec_utils import write_spec_if_changed
    write_spec_if_changed(spec_path, spec)


def write_instructions_json(test_set: str, commands: dict):
    """Write instructions/{test_set}.json from in-code instruction definitions.

    Args:
        test_set: Test set name (e.g. 'interaction_composition_x_prerequisite_chain').
        commands: Dict of {command_name: {variant_name: instruction_text}}.
    """
    instructions_dir = Path(__file__).parent.parent.parent.parent / 'instructions'
    instructions_dir.mkdir(parents=True, exist_ok=True)
    out_path = instructions_dir / f'{test_set}.json'
    data = {'commands': commands}
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
        f.write('\n')


# ═════════════════════════════════════════════════════════════════════════
# Word COM helpers
# ═════════════════════════════════════════════════════════════════════════

def open_doc(wd, path):
    """Open a document, returning the Document COM object."""
    doc = wd.Documents.Open(str(Path(path).absolute()))
    time.sleep(0.3)
    return doc


def save_as_gold(doc, gold_path):
    """Save document as gold.docx."""
    doc.SaveAs(str(Path(gold_path).absolute()), FileFormat=16)
    time.sleep(0.3)


def close_doc(doc):
    """Close document without saving."""
    for _ in range(3):
        try:
            doc.Close(False)
            break
        except Exception:
            time.sleep(1.0)
    time.sleep(0.3)
