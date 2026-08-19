"""Build IBCS templates in Excel via win32com, reproducing the SVG geometry.

The SVG renderer solves each template's geometry first; this module reproduces
it. That order matters - working out where a pin head goes is cheap in SVG and
expensive in Excel, and doing it twice guarantees the two drift.

The hard problem here is the tier stack: C03, C04, C05, C06 and C12 all need
two to four chart objects that share one category axis exactly. Excel will not
do that on its own. Autolayout sizes each plot area to fit that chart's own axis
labels, so a tier whose value labels are one character wider silently shifts its
bars sideways relative to the tier above. The fix is to stop asking:

    * give every tier chart the same Left and Width;
    * suppress the value axis on every tier, so nothing competes for width;
    * set PlotArea.InsideLeft and InsideWidth explicitly, to the same numbers,
      AFTER the series exist - Excel recomputes them when data changes;
    * set them twice, because the first assignment can be overwritten by a
      relayout triggered by the assignment itself.

verify() then measures what Excel actually did and fails loudly if the tiers do
not line up, because "looks aligned" is not a check.

The second thing to know is the zone rule. A sheet has a data zone of declared
width on the left and a chart zone to its right, and the chart zone's left edge
is measured from where the data ends rather than declared. The first version of
this module AutoFit the data block and pinned the charts at a fixed 300pt, so a
long message string in a cell pushed the data underneath them. verify() now
asserts the separation as well as the alignment.

Everything per-template - the column plan, the widths, the chart-feed formulas,
the tier geometry - lives in ibcs_layout. This module only draws what it is
handed.

Requires Windows, Excel 2016+, and pywin32.

Usage
-----
    python ibcs_excel.py --template C03A
    python ibcs_excel.py --template C03A,C04A --out "IBCS templates.xlsx"
    python ibcs_excel.py --keep-open          # leave Excel up to inspect
"""

from __future__ import annotations

import argparse
import dataclasses as dc
import shutil
import math
import sys
import tempfile
from pathlib import Path

try:
    import win32com.client as win32
except ImportError:
    win32 = None

import ibcs_data as D
import ibcs_layout as L
import ibcs_style as S
import ibcs_paths as P

# C13 is drawn by the companion `panel-charts` skill. Found rather than
# assumed, and absent rather than fatal - see ibcs_paths.
_PANEL = P.panel_charts_scripts()
if _PANEL is not None:
    sys.path.insert(0, str(_PANEL))

# Excel constants, hard-coded rather than relying on gencache having run.
XL_COLUMN_CLUSTERED = 51
XL_COLUMN_STACKED = 52
XL_BAR_CLUSTERED = 57
# 58, not 57. A waterfall is a stacked bar whose first segment is invisible,
# and reaching for the clustered constant here draws every segment from zero.
XL_BAR_STACKED = 58
XL_LINE_MARKERS = 65
XL_LINE = 4
XL_AXIS_MAXIMUM = 2
XL_CATEGORY = 1
XL_VALUE = 2
XL_NONE = -4142
XL_AUTOMATIC = -4105
XL_LABEL_OUTSIDE_END = 2
XL_LABEL_INSIDE_END = 3     # what a stacked series falls back to
XL_LABEL_CENTER = -4108     # the only position a stacked band can use
                            # that stays put when the band moves
XL_LABEL_ABOVE = 0
XL_LABEL_BELOW = 1
XL_MARKER_SQUARE = 1
XL_MARKER_CIRCLE = 8
XL_TICK_MARK_NONE = -4142
XL_TICK_LABEL_NONE = -4142
XL_TICK_LABEL_LOW = -4134
XL_OPEN_XML_WORKBOOK = 51

# Error bars, used as pin stems. Two traps here, both of which fail with the
# same useless "Exception occurred" COM error:
#   * Series.ErrorBar must be called POSITIONALLY - pywin32 does not resolve its
#     named arguments;
#   * EndStyle takes XlEndStyleCap (xlNoCap = 2), not xlNone. Passing -4142
#     poisons the chart, and every later PlotArea call then fails too, which
#     sends you hunting in the wrong place.
XL_ERROR_BAR_Y = 1
XL_ERR_INCLUDE_PLUS = 2
XL_ERR_INCLUDE_MINUS = 3
XL_ERR_TYPE_CUSTOM = -4114
XL_NO_CAP = 2

# MsoPatternType. 26 is WideUpwardDiagonal - "upward" meaning the stripes ascend
# left to right, which is what the IBCS renders use for FC.
MSO_PATTERN_WIDE_UPWARD = 26
MSO_FALSE, MSO_TRUE = 0, -1

# Layout in points. Excel positions shapes in points, the SVG works in pixels at
# 96 dpi; 0.75 converts. Keeping the same proportions as the SVG means the two
# outputs can be compared without mental arithmetic.
PT = 0.75

# Where the data zone ends, how wide each column is, and where the tiers sit are
# per-template facts, so they live in ibcs_layout rather than here. This module
# only knows how to draw whatever layout it is handed.
XL_FREE_FLOATING = 3
XL_PORTRAIT = 1
XL_LANDSCAPE = 2
MSO_COMMENT = 4
XL_BOTTOM = -4107
XL_LEFT = -4131
MSO_TEXT_HORIZONTAL = 1
# Paragraph alignment inside a text box. A column stack's tier captions are
# right-aligned into the left margin, so they end where the plot begins.
MSO_ALIGN_LEFT, MSO_ALIGN_RIGHT = 1, 3


def rgb(hex_colour: str) -> int:
    """Excel takes colours as BGR integers, not RGB - a classic silent swap."""
    h = hex_colour.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return b * 65536 + g * 256 + r


# --------------------------------------------------------------------------- #
# Formatting primitives
# --------------------------------------------------------------------------- #


def apply_scenario(fmt, scenario: str, *, point=None) -> None:
    """Put one scenario's notation onto a series or a single point.

    ``fmt`` is a ChartFormat - series.Format or point.Format - so the same code
    styles a whole series or one column of a mixed-scenario series.

    Pass ``point`` when styling a marker on a line-type series. On such a series
    Format.Line is the *connecting segment*, not the marker border, so touching
    it makes stray lines appear between adjacent pins. The marker's own border is
    MarkerForegroundColor, which is what gets set instead.
    """
    spec = S.scenario_fill(scenario)
    if spec.hatch:
        # Patterned() must be called before the colours are set; setting
        # ForeColor first and patterning after resets it to the theme colour.
        fmt.Fill.Patterned(MSO_PATTERN_WIDE_UPWARD)
        fmt.Fill.ForeColor.RGB = rgb(spec.hatch.colour)
        fmt.Fill.BackColor.RGB = rgb(spec.fill)
    else:
        fmt.Fill.Visible = MSO_TRUE
        fmt.Fill.Solid()
        fmt.Fill.ForeColor.RGB = rgb(spec.fill)

    if point is not None:
        point.MarkerForegroundColor = rgb(spec.outline or spec.fill)
        return

    if spec.outline:
        fmt.Line.Visible = MSO_TRUE
        fmt.Line.ForeColor.RGB = rgb(spec.outline)
        fmt.Line.Weight = 0.75
    else:
        fmt.Line.Visible = MSO_FALSE


def apply_variance(fmt, value: float, scenario: str, higher_is_better: bool = True) -> None:
    """Colour a variance element by impact, keeping the hatch for expected ones."""
    colour = S.variance_colour(value, higher_is_better)
    if scenario == "FC":
        fmt.Fill.Patterned(MSO_PATTERN_WIDE_UPWARD)
        fmt.Fill.ForeColor.RGB = rgb(colour)
        fmt.Fill.BackColor.RGB = rgb("#FFFFFF")
        fmt.Line.Visible = MSO_TRUE
        fmt.Line.ForeColor.RGB = rgb(colour)
        fmt.Line.Weight = 0.75
    else:
        fmt.Fill.Visible = MSO_TRUE
        fmt.Fill.Solid()
        fmt.Fill.ForeColor.RGB = rgb(colour)
        fmt.Line.Visible = MSO_FALSE


def lock_plot_area(chart, layout: L.SheetLayout, plot_extent: float,
                   cross_offset: float = 0.0, plot_lead: float = 6.0) -> None:
    """Pin the plot area so tiers cannot drift relative to one another.

    Assigned twice on purpose: the first assignment can trigger a relayout that
    overwrites it. This is the single most important call in the module.

    ``plot_extent`` is passed in rather than derived from ChartArea.Height,
    because Excel quietly resizes a chart object after creation - ask for 110pt
    and get 104pt - so anything computed from it drifts between tiers.

    Width is assigned before left, and height before top, for a reason that costs
    an hour to find otherwise: Excel validates each assignment against the
    current state of the others, so setting a left edge while the auto-sized
    width is still wide asks for a plot area that runs off the chart, and the
    call fails with a bare E_FAIL naming nothing. Shrink first, then position.

    Each pass re-fetches PlotArea rather than reusing a cached reference, and
    failures on the first pass are tolerated. Excel defers chart layout, so an
    assignment made immediately after the series are formatted can be rejected
    against geometry that is about to change; the identical call succeeds a
    moment later. The last pass is not tolerated - if the layout still will not
    take, that is a real failure and verify() must see it.
    """
    attempts = 4
    for attempt in range(attempts):
        try:
            plot = chart.PlotArea
            if layout.horizontal:
                # Transposed: the value axis is x and the shared category axis
                # is y, so it is the vertical numbers that must match across
                # tiers and the horizontal ones that are per-tier.
                plot.InsideWidth = plot_extent
                plot.InsideLeft = plot_lead
                plot.InsideHeight = layout.plot_inside_span
                plot.InsideTop = layout.plot_inside_lead + cross_offset
            else:
                plot.InsideWidth = layout.plot_inside_span
                plot.InsideLeft = layout.plot_inside_lead + cross_offset
                plot.InsideHeight = plot_extent
                plot.InsideTop = plot_lead
        except Exception:                                     # noqa: BLE001
            if attempt == attempts - 1:
                raise


def strip_chrome(chart, *, show_categories: bool) -> None:
    """Remove everything Excel adds that IBCS does not want.

    The value axis goes entirely: IBCS labels the values on the elements, and a
    value axis on a tier chart would both duplicate that and eat horizontal
    space unequally across tiers, which is what breaks alignment.
    """
    chart.HasTitle = False
    chart.HasLegend = False
    chart.ChartArea.Format.Line.Visible = MSO_FALSE
    chart.ChartArea.Format.Fill.Visible = MSO_FALSE
    chart.PlotArea.Format.Line.Visible = MSO_FALSE
    chart.PlotArea.Format.Fill.Visible = MSO_FALSE

    value_axis = chart.Axes(XL_VALUE)
    value_axis.Format.Line.Visible = MSO_FALSE
    value_axis.MajorTickMark = XL_TICK_MARK_NONE
    value_axis.MinorTickMark = XL_TICK_MARK_NONE
    value_axis.TickLabelPosition = XL_TICK_LABEL_NONE
    value_axis.HasMajorGridlines = False

    cat_axis = chart.Axes(XL_CATEGORY)
    cat_axis.MajorTickMark = XL_TICK_MARK_NONE
    cat_axis.MinorTickMark = XL_TICK_MARK_NONE
    if show_categories:
        cat_axis.TickLabelPosition = XL_TICK_LABEL_LOW
        cat_axis.TickLabels.Font.Size = 9
        cat_axis.TickLabels.Font.Name = "Arial"
    else:
        cat_axis.TickLabelPosition = XL_TICK_LABEL_NONE


def style_reference_axis(chart, reference: str) -> None:
    """The category axis of a variance tier carries the reference scenario.

    Excel can draw one line, not two, so a PL/BU double rule cannot be done on
    the axis itself - it needs a pair of drawn lines. PY's solid light rule maps
    directly, and that is what C03A needs.
    """
    spec = S.reference_axis(reference)
    axis = chart.Axes(XL_CATEGORY)
    if spec["style"] == "solid":
        axis.Format.Line.Visible = MSO_TRUE
        axis.Format.Line.ForeColor.RGB = rgb(spec["colour"])
        axis.Format.Line.Weight = spec["weight_px"] * PT
        return False

    # A double rule cannot be an axis - Excel gives an axis one line. The axis
    # is switched off entirely and draw_double_rule() puts two shapes where it
    # would have been. Returning True is how the caller knows to do that; the
    # previous version renamed the chart "[needs double rule]" and drew a single
    # line, which is the wrong notation rather than a missing one.
    axis.Format.Line.Visible = MSO_FALSE
    return True


# --------------------------------------------------------------------------- #
# Sheet
# --------------------------------------------------------------------------- #

def data_zone_width(sheet, layout: L.SheetLayout) -> float:
    """Width of the data zone in points, measured after the widths are set.

    This is what the chart zone is positioned from. The previous version pinned
    the charts at a hard-coded 300pt while calling AutoFit on the data block, so
    a long message string in a cell silently pushed the data underneath them.
    Measuring means the two zones cannot overlap by construction.
    """
    # Summed column by column, skipping the hidden ones. `Range.Width` counts
    # a hidden column's width even though nothing is drawn there, so measuring
    # the range would push the charts a third of a metre right of the data once
    # the scale block was added - a zone boundary computed from space nobody
    # can see.
    #
    # Measured on a data row, never row 1: the title rows are merged, and a
    # merged range does not report the width of the columns it spans.
    columns = getattr(layout, "columns", None)
    if columns is None:
        # The tree and panel sheets describe their data zone by width rather
        # than by a column plan, and have no engine to hide.
        return sheet.Range(sheet.Cells(layout.first_row, 1),
                           sheet.Cells(layout.first_row,
                                       layout.last_column)).Width
    total = 0.0
    for i, column in enumerate(columns, start=1):
        if i > layout.last_column:
            break
        if not column.hidden:
            total += sheet.Columns(i).Width
    return total


def write_data(sheet, t: D.Template, layout: L.SheetLayout) -> tuple[int, int]:
    """Write the data zone: two typed columns, the rest live formulas.

    Column widths are declared, never AutoFit. AutoFit sizes a column to its
    longest string, so one long cell moves the zone boundary and the charts end
    up on top of the data - which is exactly what happened.
    """
    col, first = layout.col, layout.first_row
    last = layout.last_row()

    for i, column in enumerate(layout.columns, start=1):
        sheet.Columns(i).ColumnWidth = column.width

    _write_title_block(sheet, t, layout)

    # Two-line headers, so the chart-feed columns can stay narrow and still say
    # what they are.
    header_row = layout.header_row()
    for i, text in enumerate(layout.headers(t), start=1):
        cell = sheet.Cells(header_row, i)
        cell.Value = text
        cell.Font.Bold = True
        cell.WrapText = True
        cell.VerticalAlignment = XL_BOTTOM
    sheet.Rows(header_row).RowHeight = 28.0

    reference = layout.reference(t)
    typed = [c for c in layout.columns if c.typed]

    # Formats first, values second. A statement labels its result lines
    # "= Sales revenue", and a cell that receives that string before it has been
    # told it holds text is parsed as a formula and shows #NAME? - on the chart
    # as well as in the cell, because the category axis reads the same cell.
    for i, column in enumerate(layout.columns, start=1):
        if column.number_format:
            sheet.Range(sheet.Cells(first, i),
                        sheet.Cells(last, i)).NumberFormat = column.number_format
        elif column.key in ("period", "line"):
            sheet.Range(sheet.Cells(first, i),
                        sheet.Cells(last, i)).NumberFormat = "@"

    for n, (kind, i) in enumerate(layout.sheet_entries(t)):
        row = first + n
        for key, value in layout.structure_values(t, i, kind).items():
            cell = sheet.Cells(row, col[key])
            cell.Value = value
            # A statement's components sit in from the lines they add up to.
            # IndentLevel rather than leading spaces: the spaces would be part
            # of the category label and would show up on the chart axis too.
            if (key == "line" and kind == "category" and t.rows
                    and t.rows[i].indent):
                cell.IndentLevel = t.rows[i].indent
        # Full precision, deliberately. Pre-rounding the reference or the
        # variance to one decimal is what made November read -13 where IBCS
        # printed -12: the rounding compounds with Excel's own, and the label
        # ends up a whole unit out. Let the cell hold the real number and the
        # number format do the displaying.
        for column in typed:
            cell = sheet.Cells(row, col[column.key])
            derived = layout.typed_override(t, i, column.key, row, kind)
            if derived:
                # Derived on this row even though the column is a typed one, so
                # the shading has to say so too - a shaded cell claims to be an
                # input and this one is not.
                cell.Formula = derived
                cell.Interior.ColorIndex = XL_NONE
                cell.Font.Color = rgb("#808080")
            else:
                value = _typed_value(t, layout, column.key, i, kind)
                # A summary row does not carry every typed column - a total has
                # no relative variance unless the original printed one - and an
                # empty cell is the honest answer rather than a zero, which
                # would plot as a bar of no length sitting on the axis.
                if value is not None:
                    cell.Value = value
        for key in layout.formulas:
            sheet.Cells(row, col[key]).Formula = layout.formula(key, row)
        # The scale block: a scaled copy of every drawn column for the chart to
        # read, and a text copy for its labels. Written per row like any other
        # derived column - the span cells they divide by are written once,
        # below.
        for column in layout.columns:
            if column.scaled_from or column.text_of or column.clip_of:
                sheet.Cells(row, col[column.key]).Formula = (
                    layout.scale_formula(column, row))
        # A statement's rows are not interchangeable, so three of its columns
        # depend on what kind of line the row is and cannot come from one
        # template string reused down the block.
        for key, formula in layout.row_formulas(t, i, row, kind).items():
            sheet.Cells(row, col[key]).Formula = formula

    for i, column in enumerate(layout.columns, start=1):
        cells = sheet.Range(sheet.Cells(first, i), sheet.Cells(last, i))
        # The typed cells are shaded and the derived ones greyed, so it is
        # obvious which columns are yours to change.
        if column.typed:
            cells.Interior.Color = rgb("#FFF2CC")
        elif column.key not in ("period", "scenario", "line", "sign"):
            cells.Font.Color = rgb("#808080")

    # One span per scale group, and one cell each. These are the only cells on
    # the sheet that describe the whole sheet rather than one period, which is
    # why they are not a column of formulas.
    for group in layout.scale_groups:
        sheet.Cells(first, col[f"span_{group}"]).Formula = layout.span_formula(group)

    # Hide the engine. The reader's job is two shaded columns, and burying them
    # among thirty is the opposite of the point. Hidden rather than parked
    # right, because `data_zone_width` measures the range the charts are placed
    # from and a hidden column measures zero - so the whole scale block can be
    # added without moving a single chart.
    for i, column in enumerate(layout.columns, start=1):
        if column.hidden:
            sheet.Columns(i).Hidden = True

    names = " and ".join(c.header.format(ref=reference) for c in typed)
    for column in typed:
        sheet.Cells(header_row, col[column.key]).AddComment(
            f"Typed input. {names} are the only typed columns; every other "
            f"number on this sheet is a formula and will follow.")
    return first, last


def _typed_value(t: D.Template, layout: L.SheetLayout, key: str, index: int,
                 kind: str = "category"):
    """The transcribed number behind one typed column.

    Which columns are typed differs by template - C03A transcribes the measure
    and the relative variance, C04A the measure and the absolute one - so this
    resolves a column key against the tier that actually holds that series
    rather than assuming a fixed pair.
    """
    if kind == "summary":
        row = t.summary_rows[index]
        if key in ("measure", "ac", "py"):
            return row.total if row.stack else None
        if key == "var_rel":
            return row.variance_rel
        if key == "var_abs":
            return row.variance_abs
        return None

    # A waterfall template types one column per scenario, and the tier that
    # holds it is named for the panel rather than for the column.
    for tier in t.tiers:
        if tier.kind == "waterfall" and tier.key == f"wf_{key}":
            return tier.series[0].values[index]

    measure = t.tier("measure")
    if key == "measure":
        return _pick(measure, t.category_scenarios[index], index)
    if key == "ref":
        return measure.series_for(layout.reference(t)).values[index]
    tier = t.tier(layout.typed_source.get(key, key))
    series = tier.series[0]
    return series.values[index]


def _text_row(sheet, layout: L.SheetLayout, row: int, label: str):
    """One row of the text block: a grey label in column A, the text beside it.

    The value cell is merged across the data zone on purpose. Text in an
    unmerged cell spills across its empty neighbours, and here the spill would
    run into the chart zone, where Excel draws the charts over it - so the
    message would simply vanish. Merging bounds the text to the zone that owns
    it. Merged cells do not auto-fit their row height, so heights are explicit.
    """
    name = sheet.Cells(row, 1)
    name.Value = label
    name.Font.Color = rgb("#808080")
    name.Font.Size = 9

    sheet.Range(sheet.Cells(row, 2),
                sheet.Cells(row, layout.last_column)).Merge()
    cell = sheet.Cells(row, 2)
    cell.WrapText = True
    cell.HorizontalAlignment = XL_LEFT
    cell.VerticalAlignment = XL_BOTTOM
    return cell


def _write_title_block(sheet, t: D.Template, layout: L.SheetLayout) -> None:
    """Every word the chart shows, as cells: typed inputs, then derived lines.

    This is the text half of the live-workbook rule. The numbers were put on
    formulas so a retyped figure redraws its bar; the same has to hold for the
    words, or someone with only the .xlsx can change the data and not the unit
    it is denominated in.

    What makes it work is that the unit is its own cell rather than a fragment
    of a sentence. The previous version wrote "Contribution in kEUR" as one
    string built in Python: editable, but nothing else could refer to the unit
    alone, so every tier caption would have had to repeat it as a literal and
    the sheet would have had four places to change and no single source.

    So the block is six inputs, then the lines composed from them - the subject
    line here, and one caption per variance tier, each a formula the caller then
    links a text box to.
    """
    values = {"entity": t.title.entity,
              "measure": t.title.measure,
              "unit": t.title.unit,
              "period": t.title.period,
              "reference": layout.reference(t),
              "message": t.title.message}

    for field in L.TITLE_INPUTS:
        cell = _text_row(sheet, layout, layout.title_row(field), field.capitalize())
        # Text format first: a period line of "2025" is a label, and left to
        # itself Excel stores it as the number 2025 and right-aligns it.
        cell.NumberFormat = "@"
        cell.Value = values[field]
        cell.Interior.Color = rgb("#FFF2CC")      # shaded like the typed columns

    # The derived lines. NumberFormat is left alone here - setting "@" on a cell
    # and then giving it a formula makes Excel display the formula rather than
    # evaluate it, which is a silent way to put "=B2&" in "&B3" on a chart.
    derived = [(layout.title_row("subject"), "Subject",
                f'=B{layout.title_row("measure")}&" in "&B{layout.title_row("unit")}')]
    derived += [(layout.caption_row(spec), f"{spec.key} caption",
                 layout.caption_formula(spec)) for spec in layout.captioned]

    for row, label, formula in derived:
        cell = _text_row(sheet, layout, row, label)
        cell.Formula = formula
        cell.Font.Color = rgb("#808080")          # greyed like the derived columns

    sheet.Cells(layout.title_row("subject"), 2).Font.Bold = True
    sheet.Cells(layout.title_row("subject"), 2).Font.Color = rgb("#000000")
    sheet.Cells(layout.title_row("message"), 2).Font.Size = 10

    # ~5pt per character at font size 10 is rough, but it only has to get the
    # line count right, and erring long just leaves white space. The label
    # column is not part of the space the message has to wrap into.
    zone = data_zone_width(sheet, layout) - sheet.Columns(1).Width
    wraps = max(1, -(-int(len(t.title.message) * 5.0) // int(zone - 6)))
    for row in range(1, layout.header_row()):
        sheet.Rows(row).RowHeight = 15.0
    sheet.Rows(layout.title_row("message")).RowHeight = wraps * 12.75 + 4.0


def add_chart_title(sheet, t: D.Template, layout: L.SheetLayout,
                    chart_left: float) -> list:
    """The title block the reader sees, as text boxes above the top tier.

    Linked to the title cells where Excel allows it, so the block has one source
    of truth and editing the cell updates the chart. pywin32 does not always
    accept DrawingObject.Formula, so a failure falls back to static text rather
    than aborting the build - the text is identical either way, it just stops
    tracking the cell.
    """
    sizes = {"entity": (11, False), "subject": (11, True),
             "period": (11, False), "message": (10, False)}
    shapes, top = [], layout.chart_top
    for field in L.TITLE_DISPLAY:
        size, bold = sizes[field]
        height = 26.0 if field == "message" else 15.0
        row = layout.title_row(field)
        box = _linked_textbox(sheet, layout, f"title_{layout.template_id}_{field}",
                              row, chart_left, top, layout.chart_span, height,
                              size=size, bold=bold)
        shapes.append(box)
        top += height
    return shapes


TITLE_BLOCK_HEIGHT = 71.0


def add_title_block(sheet, t: D.Template, layout, left: float, top: float,
                    width: float) -> list:
    """The title block, for a sheet whose layout is not a tier stack.

    The tier templates get theirs from `add_chart_title`, which reads geometry
    only a SheetLayout has. Six sheets - the two structure charts, the two line
    charts, the tree and the small multiples - kept their title in cells in the
    data zone instead, which meant it was never on the printed page: the print
    area covers the chart zone, and the words were on the other side of the
    sheet.

    That is not a cosmetic gap. Under IBCS the message *is* the title, so a
    printed chart without one is missing the thing it was drawn to say.

    Linked to the same input cells as everything else, so retyping the unit
    still updates every sheet at once.
    """
    sizes = {"entity": (11, False), "subject": (11, True),
             "period": (11, False), "message": (10, False)}
    shapes = []
    for field in L.TITLE_DISPLAY:
        size, bold = sizes[field]
        height = 26.0 if field == "message" else 15.0
        box = _linked_textbox(
            sheet, layout, f"title_{layout.template_id}_{field}",
            layout.title_row(field), left, top, width, height,
            size=size, bold=bold)
        shapes.append(box)
        top += height
    return shapes


def _linked_textbox(sheet, layout: L.SheetLayout, name: str, row: int,
                    left: float, top: float, width: float, height: float,
                    size: float = 11, bold: bool = False,
                    align: int = MSO_ALIGN_LEFT, wrap: bool = True):
    """A text box showing whatever a cell says, and nothing of its own.

    The link is the entire point: the text lives in the data zone, the box is
    only a view of it, and editing the cell moves the chart. pywin32 does not
    always accept DrawingObject.Formula, so a failure falls back to static text
    rather than aborting the build - the text is identical either way, it just
    stops tracking the cell.
    """
    box = sheet.Shapes.AddTextbox(MSO_TEXT_HORIZONTAL, left, top, width, height)
    box.Name = name
    box.Line.Visible = MSO_FALSE
    box.Fill.Visible = MSO_FALSE
    frame = box.TextFrame2
    frame.MarginLeft = frame.MarginRight = 0
    frame.MarginTop = frame.MarginBottom = 0
    # A title block wraps: its message is a sentence and the box is the column
    # it has to stay inside. A tier caption must not. It is two or three glyphs
    # in a box the plot geometry sized, and on a bar stack that box is whatever
    # is left to the right of the zero line - 23pt on C04A, where the caption
    # is 24pt of text. Wrapping put the "%" of "DPL%" on a second line, below a
    # 12pt box, where it simply vanished: the tier read as absolute when it was
    # relative. Unwrapped the glyph overflows a box that has no fill and no
    # border, which nobody can see.
    frame.WordWrap = MSO_TRUE if wrap else MSO_FALSE
    try:
        # The sheet name must be quoted. "=C03A!$A$1" is a valid reference
        # anywhere else in Excel but is rejected here with "Unable to set
        # the Formula property of the TextBox class", which reads like the
        # property is unavailable rather than like a syntax complaint.
        box.DrawingObject.Formula = f"='{sheet.Name}'!$B${row}"
    except Exception:                                     # noqa: BLE001
        frame.TextRange.Text = sheet.Cells(row, 2).Value
    frame.TextRange.Font.Size = size
    frame.TextRange.Font.Bold = MSO_TRUE if bold else MSO_FALSE
    frame.TextRange.Font.Name = "Arial"
    frame.TextRange.ParagraphFormat.Alignment = align
    box.Placement = XL_FREE_FLOATING
    return box


def check_caption_placement(objects: list, layout: L.SheetLayout,
                            boxes: list) -> list[str]:
    """Measure where the captions landed, rather than trusting the arithmetic.

    Two things can go wrong and neither is visible in a geometry dump: a caption
    can fall outside the tier it names, and it can sit on top of the plot and
    cover the data it is captioning. Both are checked against the tier's own
    chart object and its locked plot area.

    The margin each caption lives in is the one the plot already reserves - the
    left inset for a column stack, the headroom above the plot for a bar stack -
    so a caption that overlaps the plot means the reserve was too small, which
    is a layout fact worth failing on rather than nudging.
    """
    by_key = {f"tier_{spec.key}": obj for obj, spec in zip(objects, layout.tiers)}
    problems = []
    for box in boxes:
        obj = by_key.get("tier_" + box.Name.split("caption_", 1)[-1])
        if obj is None:
            continue
        if (box.Left < obj.Left - 0.5 or box.Top < obj.Top - 0.5
                or box.Left + box.Width > obj.Left + obj.Width + 0.5
                or box.Top + box.Height > obj.Top + obj.Height + 0.5):
            problems.append(
                f"{box.Name}: at {box.Left:.1f},{box.Top:.1f} it falls outside "
                f"{obj.Name} ({obj.Left:.1f},{obj.Top:.1f} "
                f"{obj.Width:.1f}x{obj.Height:.1f})")

        plot = obj.Chart.PlotArea
        px0, py0 = obj.Left + plot.InsideLeft, obj.Top + plot.InsideTop
        px1, py1 = px0 + plot.InsideWidth, py0 + plot.InsideHeight
        if (box.Left < px1 and box.Left + box.Width > px0
                and box.Top < py1 and box.Top + box.Height > py0):
            problems.append(
                f"{box.Name}: overlaps the plot area of {obj.Name} - it would "
                f"be drawn over the data it names")
    return problems


def add_tier_captions(sheet, layout: L.SheetLayout, objects: list,
                      chart_left: float) -> list:
    """The caption over each variance tier, linked to its cell.

    Placed from the *locked* plot geometry for the same reason the double rule
    is: a caption that says which reference scenario a tier is measured against
    has to sit on that tier's zero line, not near it, and only the locked plot
    area knows where zero fell.

    Where the caption goes is an orientation question, and both answers are the
    margin the plot already reserves. A column stack labels its tiers in the
    left margin, level with the zero line - which is where the SVG puts them,
    at ``cats.x0 - 6``. A bar stack has its tiers side by side and no left
    margin to spare, so the caption goes in the headroom above the plot, just
    right of zero.
    """
    boxes = []
    for obj, spec in zip(objects, layout.tiers):
        if not spec.caption:
            continue
        plot = obj.Chart.PlotArea
        lo, hi = spec.bounds
        fraction = (0.0 - lo) / (hi - lo)
        row = layout.caption_row(spec)
        if layout.horizontal:
            x = obj.Left + plot.InsideLeft + plot.InsideWidth * fraction
            box = _linked_textbox(
                sheet, layout, f"caption_{spec.key}", row,
                x + 3.0, obj.Top + 1.0, plot.InsideWidth * (1 - fraction) - 3.0,
                12.0, size=9, wrap=False)
        else:
            y = obj.Top + plot.InsideTop + plot.InsideHeight * (1 - fraction)
            box = _linked_textbox(
                sheet, layout, f"caption_{spec.key}", row,
                chart_left, y - 11.0, layout.plot_inside_lead - 6.0, 12.0,
                size=9, align=MSO_ALIGN_RIGHT, wrap=False)
        boxes.append(box)
    return boxes


def _pick(tier: D.Tier, scenario: str, index: int):
    series = tier.series_for(scenario)
    return None if series is None else series.values[index]


# --------------------------------------------------------------------------- #
# Charts
# --------------------------------------------------------------------------- #


def _scaled_bounds(sheet, layout: L.SheetLayout,
                   spec: L.TierSpec) -> tuple[float, float]:
    """One tier's axis bounds, as fractions of its group's span.

    Falls back to the declared bounds where the sheet has no scale block, so a
    template that has not been converted still builds and still draws exactly
    what it drew before.
    """
    lo, hi = spec.bounds
    key = f"span_{spec.scale_group}"
    if not spec.scale_group or key not in layout.col:
        return lo, hi
    span = sheet.Cells(layout.first_row, layout.col[key]).Value
    if not span:
        raise ValueError(
            f"{layout.template_id}: the {spec.scale_group} span is empty, so "
            f"every series on the {spec.key} tier would divide by nothing")
    return lo / span, hi / span


def add_tier(sheet, t: D.Template, layout: L.SheetLayout, spec: L.TierSpec,
             chart_left: float, first: int, last: int):
    """One tier as its own ChartObject, aligned to the others by construction."""
    if layout.horizontal:
        left, top = chart_left + spec.offset, layout.chart_top + layout.title_block_height
        width, height = spec.extent, layout.chart_span
    else:
        left, top = chart_left, layout.tier_offset(spec)
        width, height = layout.chart_span, spec.extent
    obj = sheet.ChartObjects().Add(left, top, width, height)
    obj.Name = f"tier_{spec.key}"
    # Excel's default is move-and-size-with-cells, which would drag the charts
    # back over the data the moment anyone widened a column. The zone split only
    # holds if the charts are anchored to the sheet, not to the cells under them.
    obj.Placement = XL_FREE_FLOATING
    chart = obj.Chart
    # The scale block is hidden, and Excel will not plot a hidden cell: a
    # series reading a hidden column comes back with no points at all, and the
    # first thing that touches Points(i) fails with "Parameter not valid"
    # rather than anything about visibility. Set before the series are added.
    chart.PlotVisibleOnly = False
    tier = t.tier(spec.key)
    if tier.kind == "waterfall" or spec.stacked:
        # Stacked, and which stacked depends on how the template runs. Both
        # earlier waterfalls were horizontal, so this was a bar constant with no
        # orientation test - and C05X, the first vertical one, came out drawn on
        # its side with the months running down the page.
        chart.ChartType = (XL_BAR_STACKED if layout.horizontal
                           else XL_COLUMN_STACKED)
    else:
        chart.ChartType = (XL_BAR_CLUSTERED if layout.horizontal
                           else XL_COLUMN_CLUSTERED)
    cats = sheet.Range(sheet.Cells(first, 1), sheet.Cells(last, 1))

    while chart.SeriesCollection().Count:
        chart.SeriesCollection(1).Delete()

    if tier.kind == "waterfall" and "wf_up_ac" in layout.col:
        _scenario_bridge_series(chart, sheet, t, layout, cats, first, last)
    elif tier.kind == "waterfall" and "wf_base" in layout.col:
        _bridge_series(chart, sheet, t, layout, cats, first, last)
    elif tier.kind == "waterfall":
        _waterfall_series(chart, sheet, t, layout, cats, first, last, spec.key)
    # Stacked first. Both of these panels feed from a bar_py column, so keying
    # the choice on the column plan alone sent C06F's stacked panel into the
    # clustered builder - which then asked Excel for an outside label on a stack
    # and got a bare COM error naming nothing. How the tier is drawn is what
    # separates them, so that is what decides.
    elif spec.key == "measure" and spec.stacked:
        _stacked_measure_series(chart, sheet, t, layout, cats, first, last)
    elif spec.key == "measure" and "bar_py" in layout.col:
        _scenario_measure_series(chart, sheet, t, layout, cats, first, last)
    elif spec.key == "measure":
        _measure_series(chart, sheet, t, layout, cats, first, last)
    elif spec.key == "var_abs":
        _variance_series(chart, sheet, t, layout, cats, first, last, fmt="+0;-0")
    else:
        _pin_series(chart, sheet, t, layout, cats, first, last)

    axis = chart.Axes(XL_VALUE)
    # The bounds are declared in data units and divided by the span the sheet
    # computed. That is the whole trick: the axis is a fixed pair of fractions,
    # the series are fractions of the same span, and any figures at all land
    # inside it. With the shipped data the span is exactly the old maximum, so
    # the picture does not move.
    axis.MinimumScale, axis.MaximumScale = _scaled_bounds(sheet, layout, spec)

    # A bar chart plots its first category at the bottom. Every table and bar
    # template reads top-down, so the order has to be reversed - and the value
    # axis crossing has to be moved with it, or the labels jump to the far side.
    if layout.horizontal:
        cat_axis = chart.Axes(XL_CATEGORY)
        cat_axis.ReversePlotOrder = True
        cat_axis.Crosses = XL_AXIS_MAXIMUM

    strip_chrome(chart, show_categories=(spec.key == layout.tiers[0].key))
    needs_rule = bool(tier.reference) and style_reference_axis(chart, tier.reference)
    lock_plot_area(chart, layout, spec.plot_extent, spec.cross_offset, spec.plot_lead)
    return obj, needs_rule


def _feed(sheet, layout: L.SheetLayout, key: str, first: int, last: int):
    """The range a chart series should read for `key`.

    The scaled copy where the sheet has a scale block, the column itself where
    it does not. Every series goes through here so that adding scaling to a
    template is a change to its column plan and not to its drawing code.
    """
    column = layout.col[layout.drawn_column(key)]
    return sheet.Range(sheet.Cells(first, column), sheet.Cells(last, column))


def _link_labels(series, sheet, layout: L.SheetLayout, key: str, first: int,
                 last: int) -> bool:
    """Point a series' data labels at the unscaled value, cell by cell.

    Only needed where the values are scaled - a label shows its series' own
    number, and a series plotting a fraction of a span would print the
    fraction. Linking is the only way back to the real figure, and the cell
    has to hold *text*: a linked label ignores the label's own number format
    and prints the cell at full precision.

    Returns False where the sheet has no scale block, so the caller can keep
    its own NumberFormat and nothing changes for a template not yet converted.
    """
    label_key = layout.label_column(key)
    if label_key is None:
        return False
    column = layout.col[label_key]
    for i in range(1, last - first + 2):
        try:
            series.Points(i).DataLabel.Formula = (
                f"='{sheet.Name}'!{a1_abs(first + i - 1, column)}")
        except Exception:                                     # noqa: BLE001
            pass
    return True


def a1_abs(row: int, column: int) -> str:
    return f"${L.col_letter(column)}${row}"


def _series_from(chart, sheet, layout, cats, first, last, key: str):
    """One series fed by one column of the data zone.

    Fed from the scaled copy where the sheet has a scale block, so every
    builder that goes through here is converted at once. The name still comes
    from the *visible* column's header, because that is what the column is
    called and the scaled copy is engine plumbing nobody reads.
    """
    series = chart.SeriesCollection().NewSeries()
    series.Values = _feed(sheet, layout, key, first, last)
    series.XValues = cats
    series.Name = sheet.Cells(layout.header_row(), layout.col[key])
    return series


def _suppress_zero_labels(series) -> None:
    """A stacked slot that does not apply on this row has length zero, and a
    label on it prints 0 against every row it does not belong to."""
    for point in range(1, series.Points().Count + 1):
        try:
            if not series.Values[point - 1]:
                series.Points(point).HasDataLabel = False
        except Exception:                                     # noqa: BLE001
            pass


def _stack_labels(series, number_format: str, sheet=None, layout=None,
                  key: str = None, first: int = None, last: int = None) -> None:
    series.HasDataLabels = True
    labels = series.DataLabels()
    for position in (XL_LABEL_OUTSIDE_END, XL_LABEL_INSIDE_END):
        try:
            labels.Position = position
            break
        except Exception:                                     # noqa: BLE001
            continue
    labels.NumberFormat = number_format
    labels.Font.Size = 9
    labels.Font.Name = "Arial"
    if key is not None:
        _link_labels(series, sheet, layout, key, first, last)
    _suppress_zero_labels(series)


def _scenario_measure_series(chart, sheet, t, layout, cats, first, last) -> None:
    """C05X's measure tier: one clustered series per scenario.

    Four fills - prior year solid light, plan outlined, actual solid dark,
    forecast hatched - and a series can carry one. Clustered rather than stacked
    so the plan sits behind the actual instead of on top of it, which is what
    the notation means; the cost is that the closing column's measured and
    expected halves stand side by side here where the original stacks them.
    That trade is recorded in references/decisions.md - a fidelity loss beats
    wrong notation.
    """
    for key, scenario in (("bar_py", "PY"), ("bar_pl", "PL"),
                          ("bar_ac", "AC"), ("bar_fc", "FC")):
        series = _series_from(chart, sheet, layout, cats, first, last, key)
        apply_scenario(series.Format, scenario)
        # Through the shared helper rather than setting the labels here, which
        # is what links them to the text column once the sheet has a scale
        # block. Hand-rolled labels print the series' own value, and a series
        # plotting a fraction of a span prints the fraction.
        _stack_labels(series, "#,##0", sheet, layout, key, first, last)


def _scenario_bridge_series(chart, sheet, t, layout, cats, first, last) -> None:
    """C05X's bridge, where impact and scenario both decide the fill.

    Four step series rather than two. A step is coloured by whether it helped or
    hurt, and filled by whether it was measured or is expected, so there are
    four combinations and a series can hold one of them. This is the densest
    piece of notation in the library: an expected improvement is hatched green,
    a measured one solid green, and the reader needs no legend to tell them
    apart.
    """
    base = _series_from(chart, sheet, layout, cats, first, last, "wf_base")
    base.Format.Fill.Visible = MSO_FALSE
    base.Format.Line.Visible = MSO_FALSE

    tier = t.tier("wf")
    plan = (("wf_up_ac", 1, "AC", "+#,##0"), ("wf_dn_ac", -1, "AC", '"-"#,##0'),
            ("wf_up_fc", 1, "FC", "+#,##0"), ("wf_dn_fc", -1, "FC", '"-"#,##0'))
    for key, sign, scenario, fmt in plan:
        series = _series_from(chart, sheet, layout, cats, first, last, key)
        # apply_variance already knows the rule: a forecast variance keeps the
        # hatch and takes its impact colour as the stripe.
        apply_variance(series.Format, sign, scenario, tier.higher_is_better)
        _stack_labels(series, fmt, sheet, layout, key, first, last)


def _stacked_measure_series(chart, sheet, t, layout, cats, first, last) -> None:
    """C06F's left panel: five fills on one axis, so five series in one slot.

    A state draws its actual solid and then appends the shortfall in prior
    year's own light grey, so the bar reaches PY and the grey is exactly what
    was lost. The three scenario totals get a series each, because PL outlined,
    PY light and AC dark are three different fills and Excel colours a series
    rather than a point - the same reason a variance is split by sign.

    Every slot holds zero where it does not apply, which is what keeps the stack
    arithmetic honest; NA() would be right for a bar that must not exist, and
    here the bar exists with no length.
    """
    plan = (("bar_state", "AC", "#,##0"),
            ("bar_short", "PY", None),
            ("bar_pl", "PL", "#,##0"),
            ("bar_py", "PY", "#,##0"),
            ("bar_ac", "AC", "#,##0"))
    for key, scenario, fmt in plan:
        series = _series_from(chart, sheet, layout, cats, first, last, key)
        apply_scenario(series.Format, scenario)
        if fmt:
            _stack_labels(series, fmt, sheet, layout, key, first, last)
        else:
            # The shortfall is not labelled. Its number is the prior year, and
            # the reference prints that only where the segment is wide enough -
            # Illinois and nowhere else. Deciding that per point would freeze it
            # at build time, so the SVG carries it and the workbook does not.
            series.HasDataLabels = False


def _bridge_series(chart, sheet, t, layout, cats, first, last) -> None:
    """C06F's bridge: a floating step per state, carrying PY onto AC.

    Three series - the invisible base, then the favourable and adverse halves of
    the step - and the split is what makes the colour follow the data rather
    than being painted on.

    The two halves hold magnitudes, because a stacked segment has a length and
    not a sign. The sign comes back through the number format: the adverse
    series prints a literal minus in front of its own value, so a step of 288
    reads -288 without anything being written onto the label. Writing text onto
    a label is what freezes it.
    """
    base = _series_from(chart, sheet, layout, cats, first, last, "wf_base")
    base.Format.Fill.Visible = MSO_FALSE
    base.Format.Line.Visible = MSO_FALSE

    tier = t.tier("wf")
    for key, sign, fmt in (("wf_up", 1, "+#,##0"), ("wf_dn", -1, '"-"#,##0')):
        series = _series_from(chart, sheet, layout, cats, first, last, key)
        series.Format.Fill.Visible = MSO_TRUE
        series.Format.Fill.Solid()
        series.Format.Fill.ForeColor.RGB = rgb(
            S.variance_colour(sign, tier.higher_is_better))
        series.Format.Line.Visible = MSO_FALSE
        _stack_labels(series, fmt, sheet, layout, key, first, last)


def _waterfall_series(chart, sheet, t, layout, cats, first, last,
                      key: str) -> None:
    """One scenario's statement as a stacked bar whose first segment is hidden.

    Three series, and each one exists for a reason worth stating.

    ``base`` is the invisible segment that floats the bar. It is a formula over
    the running level, not a column of numbers, which is the whole reason the
    workbook can be edited: retype a line and every bar below it re-floats. A
    static base looks identical on the build data and starts lying immediately.

    ``adds`` and ``subs`` are the visible length, and exactly one of them is
    non-zero on any row. That is what makes the fill follow the sign instead of
    being painted on at build time - change a line from a cost to a credit and
    the bar changes shade on recalculation. Excel colours a series, never a
    point, so this split is the only way to get it.

    Zero, not NA(), in the empty one: these are stacked segments and a stack
    needs a number in every slot. NA() is right for a bar that must not exist;
    here the bar exists and has no length.
    """
    scenario = "PY" if key.endswith("py") else "AC"
    prefix = key.split("_")[1]
    letters = layout.col

    for suffix, sign in (("base", 0), ("add", +1), ("sub", -1)):
        feed_key = f"{prefix}_{suffix}"
        column = letters[feed_key]
        series = chart.SeriesCollection().NewSeries()
        series.Values = _feed(sheet, layout, feed_key, first, last)
        series.XValues = cats
        series.Name = sheet.Cells(layout.header_row(), column)

        if sign == 0:
            series.Format.Fill.Visible = MSO_FALSE
            series.Format.Line.Visible = MSO_FALSE
            continue

        series.Format.Fill.Visible = MSO_TRUE
        series.Format.Fill.Solid()
        fill = S.waterfall_fill(scenario, sign)
        series.Format.Fill.ForeColor.RGB = rgb(fill)
        series.Format.Line.Visible = MSO_FALSE
        _waterfall_labels(series, fill)
        _link_labels(series, sheet, layout, feed_key, first, last)

    # No connectors here, and the reason is worth recording because the obvious
    # answer looks right until you draw it. Excel will join a stacked chart's
    # segments across categories - ChartGroup.HasSeriesLines - and that is the
    # documented way to build a waterfall. It joins the wrong pair of points on
    # a *floating* one: the line runs from a segment boundary in one category to
    # the same boundary in the next, and once the base moves, that boundary is
    # at a different value in each, so every connector comes out diagonal. The
    # rendered proof is in build/excel-baseline/.
    #
    # A vertical connector cannot be made from stacked series at all. It would
    # need a boundary equal to the running level in *both* the row above and the
    # row below, and each row supplies only one value per series.
    #
    # So the workbook shows the floating bars without them and the SVG carries
    # them. The alternative - drawing them as line shapes - was rejected on the
    # rule this project has now applied three times: a shape does not follow the
    # data, and a connector left at the level the statement used to reach states
    # something false, which is worse than one that is absent.


def _waterfall_labels(series, fill: str) -> None:
    """Value labels that stay live on a stacked segment.

    The label shows the series' own value, which for a waterfall segment is the
    magnitude of the line - exactly what the reference prints, including on the
    deduction rows, which it prints unsigned. So no text is ever written onto a
    label here: doing that freezes it, and a frozen label over a changed bar is
    worse than no label.

    Position is the compromise. IBCS puts the value outside the element, and
    Excel offers a stacked series only Center, InsideEnd and InsideBase - it
    rejects OutsideEnd. It is asked for anyway and the answer is taken, so if a
    later Excel allows it the chart improves without a code change.
    """
    series.HasDataLabels = True
    labels = series.DataLabels()
    for position in (XL_LABEL_OUTSIDE_END, XL_LABEL_INSIDE_END):
        try:
            labels.Position = position
            break
        except Exception:                                     # noqa: BLE001
            continue
    labels.Font.Size = 9
    labels.Font.Name = "Arial"
    labels.NumberFormat = "#,##0"
    # A stacked label sits inside its segment, so its colour is decided by the
    # fill under it rather than by preference. Luminance rather than a list of
    # dark colours: the palette is a house choice and can be changed, and a
    # hard-coded list would go stale the moment it was.
    r, g, b = (int(fill.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4))
    luminance = 0.299 * r + 0.587 * g + 0.114 * b
    labels.Font.Color = rgb("#FFFFFF" if luminance < 140 else S.TEXT["primary"])
    # A zero-length segment is the half of the pair that does not apply on this
    # row; labelling it would print a 0 against every line.
    for point in range(1, series.Points().Count + 1):
        try:
            if not series.Values[point - 1]:
                series.Points(point).HasDataLabel = False
        except Exception:                                     # noqa: BLE001
            pass


def _measure_series(chart, sheet, t, layout, cats, first, last) -> None:
    """The reference scenario behind, the measure in front, forecasts hatched.

    Two series rather than three: AC and FC live in one series and are separated
    by per-point formatting. A third series would make Excel split the category
    into three slots and shift every bar, which is exactly the kind of drift the
    tier stack cannot tolerate.
    """
    col = layout.col
    reference = layout.reference(t)

    py = chart.SeriesCollection().NewSeries()
    py.Name = reference
    py.Values = _feed(sheet, layout, "ref", first, last)
    py.XValues = cats
    apply_scenario(py.Format, reference)

    ac = chart.SeriesCollection().NewSeries()
    ac.Name = "AC"
    ac.Values = _feed(sheet, layout, "measure", first, last)
    ac.XValues = cats
    apply_scenario(ac.Format, "AC")

    for i, scenario in enumerate(t.category_scenarios, start=1):
        if scenario != "AC":
            apply_scenario(ac.Points(i).Format, scenario)

    ac.HasDataLabels = True
    labels = ac.DataLabels()
    labels.Position = XL_LABEL_OUTSIDE_END
    labels.NumberFormat = "0"
    labels.Font.Size = 9
    labels.Font.Name = "Arial"
    _link_labels(ac, sheet, layout, "measure", first, last)

    # The reference scenario sits centred in the category and the measure to
    # its right, so the comparison reads as background to this year rather than
    # as a rival bar.
    _set_cluster(chart, layout, "measure")


def _variance_series(chart, sheet, t, layout, cats, first, last, fmt: str) -> None:
    """Absolute variances as two stacked series, so colour follows the data.

    Excel colours a series, never a point-by-value. So the desirable and
    undesirable variances are plotted as two separate series, each solid in its
    own colour, with formulas making them mutually exclusive per row. Retype a
    figure so the variance flips sign and the bar changes side *and* colour on
    recalculation, with no rebuild.

    Three series in one slot, not two. The two coloured series only ever hold
    half the points each, and Excel creates a data label only where a value
    exists when labels are switched on - so a category that is empty at build
    time and gains a value after an edit draws a bar with no number against it.
    A third series carrying the *complete* variance, invisible, holds the labels:
    it has a value in every category, so every label exists from the start and
    simply follows the data. The visible colour comes from the two split series
    drawn over it.

    Clustered with Overlap at 100, not stacked. Both put the series in one slot,
    but a stacked chart refuses every outside label position - Excel offers only
    Center, InsideEnd and InsideBase - and IBCS puts value labels outside the
    element they belong to.
    """
    tier = t.tier("var_abs")
    col = layout.col
    ref = layout.reference(t)

    labeller = chart.SeriesCollection().NewSeries()
    labeller.Name = f"\u0394{ref}"
    labeller.Values = _feed(sheet, layout, "var_abs", first, last)
    labeller.XValues = cats
    labeller.Format.Fill.Visible = MSO_FALSE
    labeller.Format.Line.Visible = MSO_FALSE
    labeller.HasDataLabels = True
    marks = labeller.DataLabels()
    marks.NumberFormat = fmt
    marks.Font.Size = 9
    marks.Font.Name = "Arial"
    marks.Position = XL_LABEL_OUTSIDE_END
    _link_labels(labeller, sheet, layout, "var_abs", first, last)

    # Two series where the impact direction is fixed, four where it varies by
    # row - a statement's dPY column carries favourable and adverse bars on the
    # same side of the axis, and one series cannot hold two colours.
    for key, colour in layout.variance_splits(t, "var_abs"):
        sign = 1 if "pos" in key or key.endswith("up") else -1
        series = chart.SeriesCollection().NewSeries()
        series.Name = f"\u0394{ref} " + key.split("_", 1)[1].replace("_", " ")
        series.Values = _feed(sheet, layout, key, first, last)
        series.XValues = cats
        series.Format.Fill.Visible = MSO_TRUE
        series.Format.Fill.Solid()
        series.Format.Fill.ForeColor.RGB = rgb(colour)
        series.Format.Line.Visible = MSO_FALSE

        # Forecast variances keep the hatch. This one stays per-point and static:
        # which periods are forecast is structure, not data, so it does not move
        # when a figure is retyped.
        for i, scenario in enumerate(t.category_scenarios, start=1):
            if scenario == "FC":
                apply_variance(series.Points(i).Format, sign, scenario,
                               tier.higher_is_better)

        series.HasDataLabels = False       # the labeller series carries these

    _set_cluster(chart, layout, "var_abs")


def _set_cluster(chart, layout: L.SheetLayout, key: str) -> None:
    """Element thickness, from the layout rather than from a literal."""
    spec = next(t for t in layout.tiers if t.key == key)
    group = chart.ChartGroups(1)
    if spec.gap_width is not None:
        # Excel raises above 500, which is what rules out the obvious pin.
        group.GapWidth = min(spec.gap_width, 500)
    if spec.overlap is not None:
        group.Overlap = spec.overlap


def _pin_series(chart, sheet, t, layout, cats, first, last) -> None:
    """Relative variances as IBCS pins: a thin stem with a head marker.

    A column series cannot make the stem: GapWidth caps at 500, which on a
    twelve-category axis bottoms out around 9px where the reference uses 5px.
    So the stem is a custom Y error bar instead, whose line weight is set in
    points and therefore exact, and the head is the series' own square marker.

    Two series, because an error bar takes one colour per series and a variance
    tier needs two - desirable up in green, undesirable down in red.
    """
    if layout.horizontal:
        _pin_series_h(chart, sheet, t, layout, cats, first, last)
        return

    tier = t.tier("var_rel")
    col = layout.col
    ref = layout.reference(t)
    for name, value_key, length_key, include in (
        (f"\u0394{ref}% up", "rel_up", "rel_up_len", XL_ERR_INCLUDE_MINUS),
        (f"\u0394{ref}% down", "rel_dn", "rel_dn_len", XL_ERR_INCLUDE_PLUS),
    ):
        series = chart.SeriesCollection().NewSeries()
        series.Name = name
        series.Values = _feed(sheet, layout, value_key, first, last)
        series.XValues = cats
        series.ChartType = XL_LINE_MARKERS
        series.Format.Line.Visible = MSO_FALSE
        series.MarkerStyle = XL_MARKER_SQUARE
        series.MarkerSize = 5

        # The stem shares the head's span, because both are lengths on the
        # same axis. Scaled apart, every stem would point somewhere the head is
        # not.
        lengths = _feed(sheet, layout, length_key, first, last)
        # Positional only - see the constants block.
        series.ErrorBar(XL_ERROR_BAR_Y, include, XL_ERR_TYPE_CUSTOM, lengths, lengths)
        bars = series.ErrorBars
        bars.EndStyle = XL_NO_CAP
        bars.Format.Line.Visible = MSO_TRUE
        bars.Format.Line.Weight = 3.75            # 5px at 96dpi
        sign = 1 if include == XL_ERR_INCLUDE_MINUS else -1
        bars.Format.Line.ForeColor.RGB = rgb(
            S.variance_colour(sign, tier.higher_is_better))

        # The head carries the minuend's scenario, so a forecast variance reads
        # as forecast even in the relative tier.
        for i, scenario in enumerate(t.category_scenarios, start=1):
            apply_scenario(series.Points(i).Format, scenario, point=series.Points(i))

        series.HasDataLabels = True
        labels = series.DataLabels()
        labels.NumberFormat = "+0.0;-0.0"
        labels.Font.Size = 9
        labels.Font.Name = "Arial"
        _link_labels(series, sheet, layout, value_key, first, last)
        _place_variance_labels(series, t, "var_rel", sign)


def _pin_series_h(chart, sheet, t, layout, cats, first, last) -> None:
    """Pins on a bar chart, where the error-bar technique is not available.

    The vertical pin is a line-markers series with a custom Y error bar as the
    stem. A bar chart cannot carry one: Excel refuses to combine a bar series
    with a line series, so there is nothing to hang the error bar on and nothing
    that draws a marker.

    So the stem is a very thin bar instead - GapWidth is what makes it thin, and
    it is the reason the 500 ceiling matters. Two series again, split by sign,
    so the colour still follows the data rather than being frozen at build time.

    **Known limit.** The head marker is not drawn. On the vertical pin it is the
    series' own marker; here there is no marker to use, and drawing nineteen
    squares as shapes would make them static - which is worse than absent, since
    a head that stops tracking its pin after an edit states something false.
    Recorded in references/decisions.md with the technique that would fix it.
    """
    col = layout.col
    ref = layout.reference(t)

    labeller = chart.SeriesCollection().NewSeries()
    labeller.Name = f"Δ{ref}%"
    labeller.Values = _feed(sheet, layout, "var_rel", first, last)
    labeller.XValues = cats
    labeller.Format.Fill.Visible = MSO_FALSE
    labeller.Format.Line.Visible = MSO_FALSE
    labeller.HasDataLabels = True
    marks = labeller.DataLabels()
    marks.NumberFormat = "+0;-0"
    marks.Font.Size = 9
    marks.Font.Name = "Arial"
    marks.Position = XL_LABEL_OUTSIDE_END
    # The one series on this tier that shows a number, so the one that has to
    # be pointed back at the real figure - which the cell also decorates.
    #
    # Where a tier clips outliers the bar stops at the panel edge and the label
    # carries both the measured value and the triangles saying it was cut. One
    # label, because a separate marker series lands in exactly the same place
    # and the two print on top of each other. Getting there took three dead
    # ends worth recording: a line or scatter series *is* accepted onto a bar
    # chart and is then drawn against its own axes, arriving a row out with a
    # secondary axis in tow; custom range-driven error bars are refused
    # outright on a bar chart in every argument form; and drawing the triangles
    # as shapes works and is wrong, for the same reason the horizontal pin has
    # no head - a static ornament beside a bar that has since changed states
    # something false.
    _link_labels(labeller, sheet, layout, "var_rel", first, last)
    # A label wide enough to hold "+983" and three triangles is wider than the
    # slot Excel sizes it to, and it wraps - putting the markers on a second
    # line under the number, where they read as a separate element rather than
    # as a continuation of the bar.
    for i in range(1, last - first + 2):
        try:
            marks.Item(i).Format.TextFrame2.WordWrap = MSO_FALSE
        except Exception:                                     # noqa: BLE001
            pass

    for key, colour in layout.variance_splits(t, "var_rel"):
        series = chart.SeriesCollection().NewSeries()
        series.Name = f"Δ{ref}% " + key.split("_", 1)[1].replace("_", " ")
        series.Values = _feed(sheet, layout, key, first, last)
        series.XValues = cats
        series.Format.Fill.Visible = MSO_TRUE
        series.Format.Fill.Solid()
        series.Format.Fill.ForeColor.RGB = rgb(colour)
        series.Format.Line.Visible = MSO_FALSE
        series.HasDataLabels = False

    _set_cluster(chart, layout, "var_rel")


def draw_double_rule(sheet, obj, layout: L.SheetLayout, spec: L.TierSpec,
                     reference: str) -> list:
    """Draw a PL/BU variance axis as the two parallel rules the notation wants.

    Excel gives an axis one line. The notation needs two, because PL and BU are
    fictitious scenarios drawn outlined rather than filled, and an outlined bar
    seen edge-on is a pair of lines - which is what lets a reader tell a dPL
    tier from a dPY tier without a legend.

    So the rule is drawn as shapes over the chart rather than as the axis. The
    position is computed from the locked plot geometry and the value bounds, so
    it lands on the data zero rather than on wherever Excel put its own axis,
    and the chart's own axis line is switched off underneath.
    """
    spec_axis = S.reference_axis(reference)
    lo, hi = spec.bounds
    plot = obj.Chart.PlotArea
    shapes = []

    # Where zero actually falls inside the locked plot area.
    fraction = (0.0 - lo) / (hi - lo)
    gap = spec_axis["gap_px"] / 2 + spec_axis["weight_px"] / 2
    weight = spec_axis["weight_px"] * PT

    if layout.horizontal:
        x = obj.Left + plot.InsideLeft + plot.InsideWidth * fraction
        y0 = obj.Top + plot.InsideTop
        y1 = y0 + plot.InsideHeight
        ends = [((x + d * PT, y0), (x + d * PT, y1)) for d in (-gap, gap)]
    else:
        y = obj.Top + plot.InsideTop + plot.InsideHeight * (1 - fraction)
        x0 = obj.Left + plot.InsideLeft
        x1 = x0 + plot.InsideWidth
        ends = [((x0, y + d * PT), (x1, y + d * PT)) for d in (-gap, gap)]

    for i, ((x0, y0), (x1, y1)) in enumerate(ends):
        line = sheet.Shapes.AddLine(x0, y0, x1, y1)
        line.Name = f"rule_{spec.key}_{i}"
        line.Line.ForeColor.RGB = rgb(spec_axis["colour"])
        line.Line.Weight = weight
        line.Placement = XL_FREE_FLOATING
        shapes.append(line)
    return shapes


def _place_variance_labels(series, t: D.Template, key: str, sign: int = 0,
                           column: bool = False) -> None:
    """Labels go outside the element, in the direction it points.

    Excel's OutsideEnd puts negatives underneath the axis line rather than clear
    of the bar, so each point is positioned individually.

    Note what this no longer does. It used to overwrite label text with the value
    IBCS printed, because a derived -12.5 rendered as -13 where the original said
    -12. That was treating a symptom: the -12.5 came from pre-rounding the cell,
    and full-precision cells make Excel's own label correct. Static text would
    also defeat the point of a live workbook - retype a figure and the caption
    would keep announcing the old one.
    """
    tier = t.tier(key)
    for i, scenario in enumerate(t.category_scenarios, start=1):
        value = _pick(tier, scenario, i - 1)
        if value is None:
            continue
        if sign and (value > 0) != (sign > 0):
            continue                      # this point belongs to the other series
        # A column takes OutsideEnd, which Excel resolves correctly for both
        # signs. Above/Below are line-chart positions and are what the pin tier
        # needs; asking a column for them, or a line for OutsideEnd, is rejected.
        wanted = (XL_LABEL_OUTSIDE_END if column
                  else (XL_LABEL_ABOVE if value >= 0 else XL_LABEL_BELOW))
        try:
            series.Points(i).DataLabel.Position = wanted
        except Exception:                                     # noqa: BLE001
            pass                          # combo chart types reject some positions


# --------------------------------------------------------------------------- #
# Verification
# --------------------------------------------------------------------------- #


def verify(objects: list, expected_offsets: dict[str, float],
           layout: L.SheetLayout, shapes: list | None = None,
           template: D.Template | None = None):
    """Measure what Excel actually did. Alignment is checked, not assumed.

    The comparison is against the *intended* offset for each tier, so the pin
    tier's deliberate shift passes while an accidental drift of the same size
    would not.

    The zone check is here for the same reason. The first version of this
    workbook drew its charts on top of the data block and nothing noticed,
    because the charts themselves were correct - it was only where they sat that
    was wrong, and no assertion covered that.
    """
    problems: list[str] = []

    limit = layout.first_chart_column
    for shape in list(objects) + list(shapes or []):
        column = shape.TopLeftCell.Column
        if column < limit:
            problems.append(
                f"{shape.Name}: starts in column {L.col_letter(column)}, inside "
                f"the data zone (A:{L.col_letter(layout.last_column)}) - it "
                f"would be drawn over the source data")

    # Tiers share whichever axis they are not stacked along, so that is the one
    # measured: a column stack must agree on Left/Width and on the plot's
    # horizontal inset, a bar stack on Top/Height and the vertical one.
    geometry = []
    for obj in objects:
        plot = obj.Chart.PlotArea
        if layout.horizontal:
            box, size = obj.Top, obj.Height
            inside, inside_span = plot.InsideTop, plot.InsideHeight
        else:
            box, size = obj.Left, obj.Width
            inside, inside_span = plot.InsideLeft, plot.InsideWidth
        geometry.append((obj.Name, box, size, inside, inside_span,
                         obj.Chart.SeriesCollection().Count))

    name0, left0, width0, inside0, insidew0, _ = geometry[0]
    base0 = inside0 - expected_offsets.get(name0, 0.0)

    for name, left, width, inside, insidew, count in geometry:
        base = inside - expected_offsets.get(name, 0.0)
        if abs(left - left0) > 0.01 or abs(width - width0) > 0.01:
            problems.append(
                f"{name}: chart box {left:.2f}/{width:.2f} differs from "
                f"{name0} {left0:.2f}/{width0:.2f}")
        if abs(base - base0) > 0.5:
            problems.append(
                f"{name}: plot origin resolves to {base:.2f}pt, "
                f"{name0} to {base0:.2f}pt - tiers are not aligned")
        if abs(insidew - insidew0) > 0.5:
            problems.append(
                f"{name}: plot width {insidew:.2f}pt vs {insidew0:.2f}pt "
                f"- tiers are not aligned")
        if count == 0:
            problems.append(f"{name}: no series")

    # IBCS requires tiers measured in the same unit to be drawn at the same
    # scale. Checked rather than trusted, because getting it wrong produces a
    # chart that looks fine and overstates every variance on it.
    problems.extend(layout.check_scale())

    # The Excel caption is a formula over cells and the SVG prints a finished
    # string, so the two can drift apart while each looks right on its own.
    if template is not None:
        problems.extend(layout.check_captions(template))
    return problems, geometry


# --------------------------------------------------------------------------- #
# Entry point
# --------------------------------------------------------------------------- #


READ_ME = [
    ("IBCS® template recreations", 14, True),
    ("", 11, False),
    ("One sheet per template. Each sheet has two zones: the source data on the "
     "left, the charts to its right. They never overlap.", 11, False),
    ("", 11, False),
    ("The workbook is live. On any template sheet the two shaded columns are "
     "typed inputs - the reference scenario and the measure. Everything from "
     "the variance columns rightwards is a formula, so retyping a figure "
     "recalculates the variances and redraws the chart, including changing a "
     "bar's colour if the variance changes sign.", 11, False),
    ("", 11, False),
    ("The columns to the right of the variances feed the charts. They are "
     "derived, not data: Excel colours a whole series rather than a point, so "
     "a variance that must be green when positive and red when negative is "
     "plotted as two mutually exclusive series.", 11, False),
    ("", 11, False),
    ("The words are live too. The shaded cells at the top of each sheet hold "
     "the entity, measure, unit, period, reference scenario and message; the "
     "greyed lines under them - the subject line and each tier caption - are "
     "built from those cells, and the text on the charts is linked to them. "
     "Retype the unit once and every caption follows. Nothing a chart says is "
     "typed into the chart itself.", 11, False),
    ("", 11, False),
    ("Charts follow IBCS® notation: fill encodes how a number came to exist "
     "(solid dark actual, solid light prior year, outlined plan, hatched "
     "forecast), and variance colour encodes impact rather than sign.", 11, False),
    ("", 11, False),
    ("IBCS® does not prescribe colour codes (rule UN 4.1). The palette here "
     "is a house choice sampled to match the published reference renders.",
     10, False),
    ("", 11, False),
    ("Recreated for study from the IBCS® published templates. Not an IBCS "
     "Institute publication and not endorsed by it.", 9, False),
]


SIMPLE_NOTE = [
    ("", 11, False),
    ("This is the simple workbook.", 11, True),
    ("", 11, False),
    ("Each sheet draws the same data as the complex workbook reduced to its "
     "base tier: no variance tiers, no callouts, no highlight markers. Seven "
     "templates have a useful simple form and they are the ones here. A "
     "scattergram and a bubble chart have no tiers to drop; a line chart's "
     "tiers are all measures; and a table reduced to one column block is a "
     "list rather than a report.", 11, False),
    ("", 11, False),
    ("The data zone is identical to the complex workbook's, including the "
     "columns feeding tiers this sheet does not draw. That is deliberate - "
     "both workbooks are built from one set of numbers, so they cannot "
     "disagree with each other.", 11, False),
]


def write_read_me(sheet, simple: bool = False) -> None:
    """A cover sheet, and the workbook's own attribution.

    Its own footer, deliberately. The reference renders carry an IBCS Institute
    copyright line and a recreation must not reproduce it.
    """
    sheet.Name = "Read me"
    sheet.Columns(1).ColumnWidth = 100.0
    lines = list(READ_ME) + (SIMPLE_NOTE if simple else [])
    for row, (text, size, bold) in enumerate(lines, start=2):
        cell = sheet.Cells(row, 1)
        cell.Value = text
        cell.Font.Size = size
        cell.Font.Bold = bold
        cell.WrapText = True
        cell.VerticalAlignment = XL_BOTTOM
    sheet.Rows(2).RowHeight = 22.0


# --------------------------------------------------------------------------- #
# Table sheets
# --------------------------------------------------------------------------- #
#
# A chart sheet feeds numbers to a chart; a table sheet *is* the deliverable, so
# every rule the chart sheets follow gets stricter rather than looser.
#
# The notation survives the move intact, which is the pleasant surprise here.
# Excel already has the vocabulary: a scenario column header is a bottom border
# in the scenario's own weight and colour, and the double rule that says "plan"
# is xlDouble - the outlined fill seen edge-on, exactly as the SVG draws it.
# Nothing had to be faked with shapes.
#
# The threshold is the part worth being careful about. IBCS requires it to be
# stated wherever red is applied, so the number the conditional format tests
# against is the same cell the footnote displays: change one and the other
# cannot fail to follow, because there is only one.

XL_EDGE_TOP = 8
XL_EDGE_BOTTOM = 9
XL_CONTINUOUS = 1
XL_DOUBLE = -4119
XL_THIN = 2
XL_MEDIUM = -4138
XL_THICK = 4
XL_CELL_VALUE = 1
XL_LESS = 6
XL_LEFT_ALIGN = -4131
XL_RIGHT_ALIGN = -4152


def excel_number_format(tier: D.Tier, layout: L.TableLayout) -> str:
    """The Excel format for a column, derived from the one the SVG prints with.

    Both come from ``Tier.number_format`` so a column cannot show one precision
    on the page and another in the workbook - the cumulative dPY prints to one
    decimal in both, because the source does.
    """
    decimals = 0
    if "." in tier.number_format:
        decimals = int(tier.number_format.split(".")[1][0])
    if tier.kind == "measure":
        return layout.measure_format.format(
            d="" if not decimals else "." + "0" * decimals)
    body = "0" if not decimals else "0." + "0" * decimals
    suffix = '"%"' if tier.kind == "variance_rel" else ""
    # Conditional sections, because Excel's plain third section fires only on an
    # exact zero while the convention is about a value that *rounds* to zero:
    # the cumulative ΔPY% of +0.0215 is printed "0.0%", not "+0.0%". This is the
    # same rule Tier.label_for applies in the SVG, said in Excel's own language.
    edge = 0.5 / (10 ** decimals)
    return (f"[>{edge:g}]+{body}{suffix};[<-{edge:g}]-{body}{suffix};"
            f"{body}{suffix}")


def _write_input_block(sheet, t: D.Template, layout, reference: str) -> None:
    """The six typed inputs and the subject line built from them.

    Every sheet that is not a chart sheet writes this same block, so it is
    written once. A chart sheet needs more - it has to link text boxes to these
    cells - which is why ``_write_title_block`` is still its own thing; the
    tables, the structure panels and the lines all want exactly this and nothing
    more.
    """
    values = {"entity": t.title.entity, "measure": t.title.measure,
              "unit": t.title.unit, "period": t.title.period,
              "reference": reference, "message": t.title.message}
    for field in L.TITLE_INPUTS:
        row = layout.title_row(field)
        label = sheet.Cells(row, 1)
        label.Value = field.capitalize()
        label.Font.Color = rgb("#808080")
        cell = sheet.Cells(row, 2)
        cell.NumberFormat = "@"
        cell.Value = values[field]
        cell.Interior.Color = rgb("#FFF2CC")      # shaded like every typed input

    subject = sheet.Cells(layout.title_row("subject"), 1)
    subject.Value = "Subject"
    subject.Font.Color = rgb("#808080")
    cell = sheet.Cells(layout.title_row("subject"), 2)
    # A template with no single unit leaves it blank, and the subject line has
    # to stop rather than trail off in a dangling "in". C10 plots three measures
    # in three units, so none of them can head the chart.
    cell.Formula = (f'=B{layout.title_row("measure")}'
                    f'&IF(B{layout.title_row("unit")}="",""," in "'
                    f'&B{layout.title_row("unit")})')
    cell.Font.Color = rgb("#808080")


def _write_table_title(sheet, t: D.Template, layout: L.TableLayout) -> None:
    """The typed inputs, then the header the reader actually reads.

    Same split as a chart sheet - inputs above, presentation below - but the
    presentation is cells rather than text boxes, so the three header lines are
    plain formulas over the inputs. Retype the unit in B3 and the subject line
    follows; retype the period and both block headers follow, because they are
    formulas over it too.
    """
    _write_input_block(sheet, t, layout, t.tiers[1].reference or "")

    for row, formula, size, bold in (
            (layout.entity_row, f'=B{layout.title_row("entity")}', 11, False),
            (layout.subject_row, f'=B{layout.title_rows}', 12, True),
            (layout.period_row, f'=B{layout.title_row("period")}', 11, False)):
        cell = sheet.Cells(row, 1)
        cell.Formula = formula
        cell.Font.Size = size
        cell.Font.Bold = bold


def _write_table_header(sheet, t: D.Template, layout: L.TableLayout) -> None:
    """Block headers, column captions, and the rule that says what each holds."""
    plan = layout.plan(t)
    blocks: dict[str, list[int]] = {}

    for i, entry in enumerate(plan, start=1):
        if entry.kind == "label":
            cell = sheet.Cells(layout.caption_row, i)
            cell.Borders(XL_EDGE_BOTTOM).LineStyle = XL_CONTINUOUS
            cell.Borders(XL_EDGE_BOTTOM).Weight = XL_THIN
            continue
        blocks.setdefault(entry.tier.block, []).append(i)

        caption = sheet.Cells(layout.caption_row, i)
        if entry.caption:
            caption.Value = entry.caption
        elif entry.tier.kind != "measure":
            # The unlabelled half of a pair. The heading belongs over the end
            # of the pair, as the reference prints it, so it moves here.
            previous = plan[i - 2]
            if (previous.kind == "value"
                    and previous.tier.reference == entry.tier.reference
                    and previous.caption):
                caption.Value = previous.caption
                sheet.Cells(layout.caption_row, i - 1).Value = ""
        caption.HorizontalAlignment = XL_RIGHT_ALIGN
        border = caption.Borders(XL_EDGE_BOTTOM)
        if entry.tier.kind == "measure":
            _scenario_border(border, entry.series.scenario)
        else:
            border.LineStyle = XL_CONTINUOUS
            border.Weight = XL_THIN
            border.Color = rgb(S.TEXT["primary"])

    # The block header is a formula over the period input: the month name, and
    # the same name under the underscore prefix that UN 4.2 uses for a year to
    # date. Retyping "November 2026" as "December 2026" moves both.
    period = f'B{layout.title_row("period")}'
    month = f'IF(ISERROR(FIND(" ",{period})),{period},LEFT({period},FIND(" ",{period})-1))'
    for n, (block, indices) in enumerate(blocks.items()):
        if not block:
            continue        # a single-block table has nothing to head
        cell = sheet.Cells(layout.block_header_row, indices[0])
        cell.Formula = f'={month}' if not block.startswith("_") else f'="_"&{month}'
        cell.HorizontalAlignment = XL_LEFT_ALIGN
        rule = sheet.Range(sheet.Cells(layout.block_header_row, indices[0]),
                           sheet.Cells(layout.block_header_row, indices[-1]))
        rule.Borders(XL_EDGE_BOTTOM).LineStyle = XL_CONTINUOUS
        rule.Borders(XL_EDGE_BOTTOM).Weight = XL_THIN
        rule.Borders(XL_EDGE_BOTTOM).Color = rgb(S.TEXT["primary"])


def _scenario_border(border, scenario: str) -> None:
    """A column header rule in its scenario's notation.

    The same three statements the fills make: solid dark for an actual, solid
    light for a prior year, and - the one Excel gets right for free - a double
    rule for the outlined plan.
    """
    spec = S.scenario_fill(scenario)
    if spec.is_fictitious:
        border.LineStyle = XL_DOUBLE
        border.Weight = XL_THICK
        border.Color = rgb(spec.outline)
    else:
        border.LineStyle = XL_CONTINUOUS
        border.Weight = XL_THICK
        border.Color = rgb(spec.fill)


def _write_table_body(sheet, t: D.Template, layout: L.TableLayout) -> None:
    """Row labels, the typed scenario values, and a formula for everything else."""
    plan = layout.plan(t)
    columns = layout.columns(t)
    measure = layout.measure_columns(t)

    for i, row in enumerate(t.rows):
        r = layout.first_row + i
        sheet.Rows(r).RowHeight = layout.row_height

        label = sheet.Cells(r, columns["label"])
        # Text format *before* the value. A statement labels its result lines
        # "= Sales revenue", and a cell that receives that string before it has
        # been told it holds text is parsed as a formula and shows #NAME?.
        label.NumberFormat = "@"
        label.Value = f"{row.prefix} {row.label}".strip() if row.prefix else row.label
        label.Font.Bold = row.kind == "subtotal"
        label.Font.Italic = row.kind == "ratio"
        label.IndentLevel = row.indent + (0 if row.prefix else 1)

        for n, entry in enumerate(plan, start=1):
            if entry.kind == "label":
                continue
            tier = entry.tier
            cell = sheet.Cells(r, n)
            cell.NumberFormat = (layout.ratio_formats[tier.kind]
                                 if row.kind == "ratio" and tier.kind in layout.ratio_formats
                                 else excel_number_format(tier, layout))
            cell.Font.Bold = row.kind == "subtotal"
            cell.Font.Italic = row.kind == "ratio"

            if tier.key in t.panel_tiers:
                # The panel draws this figure, so the cell must not also print
                # it under a transparent chart. The formula stays: click the
                # cell and the number is in the formula bar, which is where
                # someone tracing a bar back to its arithmetic will look.
                cell.NumberFormat = ";;;"
            if tier.kind == "measure":
                letter = L.col_letter(n)
                if row.kind == "ratio":
                    # A margin is a ratio of two rows of its own column, so it
                    # is a formula like any other total - retype a revenue line
                    # and the margin follows it.
                    top = layout.first_row + t.ratio_of[0]
                    bottom = layout.first_row + t.ratio_of[1]
                    cell.Formula = (f"=IF({letter}{bottom}=0,NA(),"
                                    f"{letter}{top}/{letter}{bottom}*100)")
                elif row.kind == "subtotal":
                    # A total with components is a formula. The same rule the
                    # waterfalls learned the hard way, and the reason a retyped
                    # country moves its continent and the world.
                    cell.Formula = L.subtotal_formula(t, layout, i, letter)
                else:
                    cell.Value = entry.value(i)
                    cell.Interior.Color = rgb("#FFF2CC")
            elif entry.value(i) is None:
                # The template has no value here - T04A's margin row carries no
                # variance - so the cell stays empty rather than holding a
                # number nothing draws.
                pass
            else:
                ac = measure[(tier.block, "AC")]
                ref = measure[(tier.block, tier.reference)]
                if tier.kind == "variance_abs":
                    cell.Formula = f"={ac}{r}-{ref}{r}"
                else:
                    cell.Formula = (f"=IF({ref}{r}=0,NA(),"
                                    f"({ac}{r}-{ref}{r})/{ref}{r}*100)")

        _row_rules(sheet, t, layout, i, len(plan))


def _row_rules(sheet, t: D.Template, layout: L.TableLayout, index: int,
               width: int) -> None:
    """The light rule under a row, and the heavier one that closes a group."""
    row = t.rows[index]
    r = layout.first_row + index
    span = sheet.Range(sheet.Cells(r, 1), sheet.Cells(r, width))

    if row.kind == "subtotal":
        top = span.Borders(XL_EDGE_TOP)
        top.LineStyle = XL_CONTINUOUS
        top.Weight = XL_MEDIUM
        top.Color = rgb(S.TEXT["primary"])
    elif index + 1 < len(t.rows) and t.rows[index + 1].kind == "subtotal":
        return                       # the subtotal's own rule closes this row
    else:
        bottom = span.Borders(XL_EDGE_BOTTOM)
        bottom.LineStyle = XL_CONTINUOUS
        bottom.Weight = XL_THIN
        bottom.Color = rgb("#BDBDBD")


def _write_table_footnote(sheet, t: D.Template, layout: L.TableLayout) -> list:
    """The threshold under each variance column, as the cell the red reads.

    Returns the addresses so the conditional formats can point at them. This is
    the whole reason the footnote is a cell rather than a caption: IBCS wants
    the threshold stated wherever it is applied, and a stated number that the
    application does not use is worse than no number at all.
    """
    plan = layout.plan(t)
    row = layout.footnote_row(t)
    addresses = {}

    first_marked = next((n for n, e in enumerate(plan, start=1)
                         if e.kind == "value" and e.tier.red_below is not None),
                        None)
    if first_marked is None:
        return addresses
    lead = sheet.Cells(row, max(1, first_marked - 1))
    lead.Value = "Δ in red if:"
    lead.Font.Italic = True
    lead.Font.Color = rgb(S.PAGE["footnote"])
    lead.HorizontalAlignment = XL_RIGHT_ALIGN

    for n, entry in enumerate(plan, start=1):
        if entry.kind == "label" or entry.tier.red_below is None:
            continue
        cell = sheet.Cells(row, n)
        cell.NumberFormat = (layout.threshold_format_rel
                             if entry.tier.kind == "variance_rel"
                             else layout.threshold_format_abs)
        cell.Value = entry.tier.red_below
        cell.Font.Italic = True
        cell.Font.Color = rgb(S.PAGE["footnote"])
        addresses[n] = f"${L.col_letter(n)}${row}"
    return addresses


def _apply_table_red(sheet, t: D.Template, layout: L.TableLayout,
                     thresholds: dict) -> None:
    """Red past the stated threshold - and only ever the stated threshold."""
    for n, address in thresholds.items():
        rng = sheet.Range(sheet.Cells(layout.first_row, n),
                          sheet.Cells(layout.last_row(t), n))
        rng.FormatConditions.Delete()
        condition = rng.FormatConditions.Add(Type=XL_CELL_VALUE,
                                             Operator=XL_LESS,
                                             Formula1=f"={address}")
        condition.Font.Color = rgb(S.VARIANCE["bad"])


def workbook_values(t: D.Template) -> dict:
    """What the sheet's own formulas should produce, computed the same way.

    Deliberately *not* the transcribed values. The workbook derives its
    subtotals from the lines rather than taking the printed ones, so on a
    statement printed to one decimal the two legitimately differ by a rounding -
    T03A's result before tax walks to 418.1 where the reference prints 418.2 -
    and every percentage taken from it differs in turn.

    Checking the sheet against the transcription would flag that as an error on
    every build. Checking it against this tests what is actually in question:
    whether the formulas do what they claim. That the transcription matches the
    reference is check_ties' job, and a different question.

    Returns {(tier key, series index, row): value}.
    """
    measures: dict[tuple[str, int], list] = {}
    for tier in t.tiers:
        if tier.kind != "measure":
            continue
        for index, series in enumerate(tier.series):
            values = list(series.values)
            spans = D.waterfall_spans(t, values)
            for i, row in enumerate(t.rows):
                if row.kind == "subtotal":
                    lo, hi = spans[i]
                    values[i] = hi if row.spans == "zero" else abs(hi - lo)
            if t.ratio_of:
                top, bottom = t.ratio_of
                for i, row in enumerate(t.rows):
                    if row.kind == "ratio":
                        values[i] = values[top] / values[bottom] * 100
            measures[(tier.key, index)] = values

    expected = {(key[0], key[1], i): v
                for key, values in measures.items()
                for i, v in enumerate(values)}

    for tier in t.tiers:
        if tier.kind == "measure":
            continue
        block_measure = next(m for m in t.tiers
                             if m.kind == "measure" and m.block == tier.block)
        actual = measures[(block_measure.key,
                           [s.scenario for s in block_measure.series].index("AC"))]
        reference = measures[(block_measure.key,
                              [s.scenario for s in block_measure.series]
                              .index(tier.reference))]
        for i in range(len(t.rows)):
            if tier.series[0].values[i] is None:
                expected[(tier.key, 0, i)] = None
                continue
            gap = actual[i] - reference[i]
            expected[(tier.key, 0, i)] = (
                gap if tier.kind == "variance_abs" else gap / reference[i] * 100)
    return expected


def verify_table(sheet, t: D.Template, layout: L.TableLayout) -> list[str]:
    """Prove the sheet says what the data says, before anyone trusts it.

    Three claims, each of which has caught something: that every subtotal cell
    holds a formula rather than a number, that what Excel computed equals what
    check_ties proved, and that every cell the reference prints in red is red
    on the sheet - read back from Excel's own evaluation of the condition, not
    from our arithmetic repeated.
    """
    problems: list[str] = []
    plan = layout.plan(t)
    wanted = workbook_values(t)

    for n, entry in enumerate(plan, start=1):
        if entry.kind == "label":
            continue
        for i, row in enumerate(t.rows):
            cell = sheet.Cells(layout.first_row + i, n)
            expected = wanted[(entry.tier.key, entry.series_index, i)]
            if expected is None:
                if str(cell.Formula).strip():
                    problems.append(
                        f"{layout.template_id}: {row.label} {entry.caption} "
                        f"holds {cell.Formula!r} where the template has no value")
                continue
            derived = (entry.tier.kind != "measure"
                       or row.kind in ("subtotal", "ratio"))
            if derived and not str(cell.Formula).startswith("="):
                problems.append(
                    f"{layout.template_id}: {row.label} {entry.caption} is typed "
                    f"({cell.Formula!r}) where it should be a formula")
            got = cell.Value
            if got is None or abs(float(got) - float(expected)) > 0.005:
                problems.append(
                    f"{layout.template_id}: {row.label} {entry.caption} shows "
                    f"{got} where the data says {expected:.3f}")

    reference_red = {(n, i)
                     for n, entry in enumerate(plan, start=1)
                     if entry.kind == "value" and entry.tier.red_below is not None
                     for i in range(len(t.rows))
                     if entry.value(i) < entry.tier.red_below}
    on_sheet = set()
    for n, entry in enumerate(plan, start=1):
        if entry.kind == "label" or entry.tier.red_below is None:
            continue
        for i in range(len(t.rows)):
            cell = sheet.Cells(layout.first_row + i, n)
            if cell.DisplayFormat.Font.Color == rgb(S.VARIANCE["bad"]):
                on_sheet.add((n, i))
    if on_sheet != reference_red:
        problems.append(
            f"{layout.template_id}: the red cells disagree with the threshold - "
            f"only on the sheet {sorted(on_sheet - reference_red)}, "
            f"only in the data {sorted(reference_red - on_sheet)}")
    return problems


def _panel_columns_of(t: D.Template, layout: L.TableLayout) -> list:
    """(excel column, plan entry) for each drawn column, in sheet order."""
    return [(n, e) for n, e in enumerate(layout.plan(t), start=1)
            if e.kind == "value" and e.tier.key in t.panel_tiers]


def _write_panel_scaffolding(sheet, t: D.Template, layout: L.TableLayout) -> dict:
    """The split columns that let a bar change colour when its sign changes.

    Excel colours a series, never a point, so a variance that can be either
    favourable or adverse needs two series with formulas making them mutually
    exclusive. Kept beside the table rather than hidden: it is how the panels are
    fed, and someone tracing a bar back to a number should be able to get there.
    """
    plan = layout.plan(t)
    start = len(plan) + 2
    scaffolding: dict[str, tuple[int, int]] = {}

    header = sheet.Cells(layout.caption_row, start)
    header.Value = "chart feed"
    header.Font.Italic = True
    header.Font.Color = rgb(S.PAGE["footnote"])

    n = start
    last = layout.last_row(t)

    # One span per scale group, ahead of the columns that divide by it. Same
    # trick as the chart sheets: Excel will not bind an axis bound to a
    # formula, so the data is divided by a span instead and the bounds become
    # fractions of it. Panels sharing a group divide by one cell, which is what
    # keeps a month's dPL panel on the same ruler as the year's.
    spans: dict[str, str] = {}
    groups: dict[str, list[str]] = {}
    for column, entry in _panel_columns_of(t, layout):
        geometry = layout.panels[entry.tier.key]
        if geometry.scale_group:
            groups.setdefault(geometry.scale_group, []).append(
                f"{L.col_letter(column)}{layout.first_row}:"
                f"{L.col_letter(column)}{last}")
    for group, ranges in groups.items():
        sheet.Columns(n).ColumnWidth = 9.0
        cap = sheet.Cells(layout.caption_row, n)
        cap.Value = f"span {group}"
        cap.Font.Size = 8
        cap.Font.Color = rgb(S.PAGE["footnote"])
        span = ",".join(ranges)
        # AGGREGATE, because a variance column is blank on the rows it does not
        # apply to and MAX over an error is an error. Option 6 ignores them.
        sheet.Cells(layout.first_row, n).Formula = (
            f"=MAX(AGGREGATE(4,6,{span}),-AGGREGATE(5,6,{span}),1E-9)")
        sheet.Cells(layout.first_row, n).NumberFormat = "0.000000"
        spans[group] = f"${L.col_letter(n)}${layout.first_row}"
        n += 1

    for column, entry in _panel_columns_of(t, layout):
        letter = L.col_letter(column)
        geometry = layout.panels[entry.tier.key]
        divide = f"/{spans[geometry.scale_group]}" if geometry.scale_group else ""

        # The invisible labeller plots the value and shows the number. Once the
        # bars are fractions it has to be a fraction too, or its labels land at
        # the wrong end of the panel - and then the number it prints has to come
        # from a cell, because a linked label ignores its own number format.
        sheet.Columns(n).ColumnWidth = 7.0
        cap = sheet.Cells(layout.caption_row, n)
        cap.Value = f"{entry.caption} scaled"
        cap.Font.Size = 8
        cap.Font.Color = rgb(S.PAGE["footnote"])
        for i in range(len(t.rows)):
            r = layout.first_row + i
            sheet.Cells(r, n).Formula = (
                f"=IF(ISBLANK({letter}{r}),NA(),{letter}{r}{divide})")
        scaled_col = n
        n += 1

        sheet.Columns(n).ColumnWidth = 7.0
        cap = sheet.Cells(layout.caption_row, n)
        cap.Value = f"{entry.caption} text"
        cap.Font.Size = 8
        cap.Font.Color = rgb(S.PAGE["footnote"])
        fmt = excel_number_format(entry.tier, layout).replace('"', '""')
        for i in range(len(t.rows)):
            r = layout.first_row + i
            sheet.Cells(r, n).Formula = (
                f'=IF(ISBLANK({letter}{r}),"",TEXT({letter}{r},"{fmt}"))')
        text_col = n
        n += 1

        # Split by *impact*, not by sign. On a statement the two are not the
        # same thing: a cost line up is adverse and a cost line down is
        # favourable, so within one column the favourable bars sit on both
        # sides of the axis. Splitting by sign would paint every increase one
        # colour, which is the error IBCS exists to prevent.
        for suffix in ("good", "bad"):
            sheet.Columns(n).ColumnWidth = 7.0
            caption = sheet.Cells(layout.caption_row, n)
            caption.Value = f"{entry.caption} {suffix}"
            caption.Font.Size = 8
            caption.Font.Color = rgb(S.PAGE["footnote"])
            for i, row in enumerate(t.rows):
                r = layout.first_row + i
                desirable = ">0" if t.higher_is_better_at(i, entry.tier) else "<0"
                adverse = "<0" if desirable == ">0" else ">0"
                test = desirable if suffix == "good" else adverse
                # The test stays on the *raw* value. Dividing by a positive
                # span cannot change a sign, but testing the fraction would
                # make the colour depend on the scale block, and impact is a
                # property of the figure rather than of how it is drawn.
                sheet.Cells(r, n).Formula = (
                    f"=IF(ISBLANK({letter}{r}),NA(),"
                    f"IF({letter}{r}{test},{letter}{r}{divide},NA()))")
            n += 1
        scaffolding[entry.tier.key] = (n - 2, n - 1, scaled_col, text_col,
                                       spans.get(geometry.scale_group))
    return scaffolding


def _add_panel_charts(sheet, t: D.Template, layout: L.TableLayout,
                      scaffolding: dict) -> list:
    """One chart per drawn column, anchored to the rows it describes.

    The alignment problem a table poses that a chart sheet does not: the bars
    have to line up with *cell* boundaries, not with a plot area of our own
    choosing. It comes out exact because the chart is positioned from the cells
    themselves - Cells(...).Top and .Left, never arithmetic over row heights -
    and because the plot area is then locked to fill the chart, so twenty
    categories divide the same height as twenty rows.

    It also only works because the sheet runs its rows contiguously. The
    reference leaves a gap between continents; reproducing that would put blank
    categories in every panel to keep the two in step, and a blank category is a
    row the chart draws nothing on rather than a gap between rows.
    """
    objects = []
    last = layout.last_row(t)

    for column, entry in _panel_columns_of(t, layout):
        geometry = layout.panels[entry.tier.key]
        anchor = sheet.Cells(layout.first_row, column)
        obj = sheet.ChartObjects().Add(
            anchor.Left, anchor.Top, sheet.Columns(column).Width,
            sheet.Cells(last, column).Top + sheet.Cells(last, column).Height
            - anchor.Top)
        obj.Name = f"panel_{entry.tier.key}"
        chart = obj.Chart
        chart.ChartType = XL_BAR_CLUSTERED

        cats = sheet.Range(sheet.Cells(layout.first_row, layout.columns(t)["label"]),
                           sheet.Cells(last, layout.columns(t)["label"]))
        values = sheet.Range(sheet.Cells(layout.first_row, column),
                             sheet.Cells(last, column))

        # The labeller: invisible, complete, and the only series carrying data
        # labels. The two coloured series each hold half the rows, so a label
        # attached to them would not exist for a row that changes sign later.
        good, bad, scaled, text_col, span_ref = scaffolding[entry.tier.key]

        labeller = chart.SeriesCollection().NewSeries()
        labeller.Name = entry.caption
        labeller.Values = sheet.Range(sheet.Cells(layout.first_row, scaled),
                                      sheet.Cells(last, scaled))
        labeller.XValues = cats
        labeller.Format.Fill.Visible = MSO_FALSE
        labeller.Format.Line.Visible = MSO_FALSE
        labeller.HasDataLabels = True
        marks = labeller.DataLabels()
        marks.NumberFormat = excel_number_format(entry.tier, layout)
        marks.Font.Size = 9
        marks.Font.Name = "Arial"
        marks.Position = XL_LABEL_OUTSIDE_END
        # The label reads a cell rather than the series, because the series now
        # holds a fraction of the span. The cell holds text, because a linked
        # label ignores its own number format and prints full precision.
        for i in range(len(t.rows)):
            try:
                labeller.Points(i + 1).DataLabel.Formula = (
                    f"='{sheet.Name}'!{a1_abs(layout.first_row + i, text_col)}")
            except Exception:                                 # noqa: BLE001
                pass
        # A panel is narrow and Excel wraps a label to fit it, so "+47" comes
        # out as "+4" over "7". There is no width to give it - the panel width
        # is the scale - so the wrapping is what has to go.
        for i in range(1, len(t.rows) + 1):
            try:
                marks.Item(i).Format.TextFrame2.WordWrap = MSO_FALSE
            except Exception:                                 # noqa: BLE001
                pass

        for source, colour in ((good, S.VARIANCE["good"]),
                               (bad, S.VARIANCE["bad"])):
            series = chart.SeriesCollection().NewSeries()
            series.Name = f"{entry.caption} {'good' if source == good else 'bad'}"
            series.Values = sheet.Range(sheet.Cells(layout.first_row, source),
                                        sheet.Cells(last, source))
            series.XValues = cats
            series.Format.Fill.Visible = MSO_TRUE
            series.Format.Fill.Solid()
            series.Format.Fill.ForeColor.RGB = rgb(colour)
            series.Format.Line.Visible = MSO_FALSE
            series.HasDataLabels = False

        # A pin is a bar made thin, which is all GapWidth is for. 500 is Excel's
        # ceiling and gives a sixth of the row band - the stem width the SVG
        # draws. The head marker cannot be drawn on a bar chart at all; that
        # limit is C04A's and is recorded in references/decisions.md.
        chart.ChartGroups(1).GapWidth = (
            500 if entry.tier.kind == "variance_rel" else 118)
        chart.ChartGroups(1).Overlap = 100

        # Fractions of the span, for the same reason the tier charts use them:
        # the axis is fixed at build and the data moves to meet it. With the
        # shipped figures the span is what the old bounds were measured
        # against, so the picture does not move.
        low, high = geometry.bounds
        if span_ref:
            span = sheet.Range(span_ref).Value
            if not span:
                raise ValueError(
                    f"{layout.template_id}: the {geometry.scale_group} span is "
                    f"empty, so every panel in it would divide by nothing")
            low, high = low / span, high / span
        value_axis = chart.Axes(XL_VALUE)
        value_axis.MinimumScale = low
        value_axis.MaximumScale = high
        chart.Axes(XL_CATEGORY).ReversePlotOrder = True

        strip_chrome(chart, show_categories=False)
        # A bar chart draws its category axis *at zero*, which is exactly where
        # the double rule goes - so left on, Excel contributes a third line to a
        # two-line notation. The rule is drawn as shapes; the axis is not needed.
        chart.Axes(XL_CATEGORY).Format.Line.Visible = MSO_FALSE
        # Transparent. The plot is exactly the width of its column, but the
        # chart around it is wider - Excel keeps a margin - so an opaque chart
        # would white out the neighbouring column. Nothing is drawn in that
        # margin, so letting it through costs nothing.
        chart.ChartArea.Format.Fill.Visible = MSO_FALSE
        chart.ChartArea.Format.Line.Visible = MSO_FALSE

        _fit_panel_plot(obj, anchor.Left, anchor.Top,
                        sheet.Columns(column).Width,
                        sheet.Cells(last, column).Top
                        + sheet.Cells(last, column).Height - anchor.Top)
        # Only now: the rule lands on the data zero inside the *fitted* plot,
        # so it has to be drawn after the plot has stopped moving.
        _panel_zero_rule(sheet, obj, entry, geometry)
        objects.append(obj)

    return objects


def _panel_zero_rule(sheet, obj, entry, geometry: L.PanelGeometry) -> list:
    """The panel's zero, drawn in the notation of the scenario it measures against.

    The vertical form of the same statement C04A's variance tier makes along the
    bottom: two parallel rules for a plan comparison, because plan is drawn
    outlined and an outlined bar seen edge-on is a pair of lines. Excel gives an
    axis one line, so the pair is drawn as shapes over the chart and the axis
    itself is switched off underneath.

    The position comes from the plot geometry and the value bounds rather than
    from where Excel put its axis, which is the same reason draw_double_rule
    does it that way.
    """
    axis = S.reference_axis(entry.tier.reference)
    plot = obj.Chart.PlotArea
    x = obj.Left + plot.InsideLeft + plot.InsideWidth * geometry.zero_fraction
    y0 = obj.Top + plot.InsideTop
    y1 = y0 + plot.InsideHeight

    offsets = ((0.0,) if axis["style"] == "solid"
               else (-(axis["gap_px"] / 2 + axis["weight_px"] / 2),
                     +(axis["gap_px"] / 2 + axis["weight_px"] / 2)))
    shapes = []
    for i, d in enumerate(offsets):
        line = sheet.Shapes.AddLine(x + d * PT, y0, x + d * PT, y1)
        line.Name = f"zero_{entry.tier.key}_{i}"
        line.Line.ForeColor.RGB = rgb(axis["colour"])
        line.Line.Weight = axis["weight_px"] * PT
        line.Placement = XL_FREE_FLOATING
        shapes.append(line)
    return shapes


def _lock_panel_plot(chart, width: float, height: float) -> None:
    """Make the plot area fill the chart, so a category band equals a row.

    Same two rules as lock_plot_area, learned the same way: assign the size
    before the position - Excel validates each assignment against the others,
    so moving a left edge while the auto-sized width is still wide asks for a
    plot area hanging off the chart and fails with a bare E_FAIL - and retry,
    because chart layout is deferred and the identical call succeeds a moment
    later.
    """
    for attempt in range(4):
        try:
            plot = chart.PlotArea
            plot.InsideWidth = width
            plot.InsideLeft = 0
            plot.InsideHeight = height
            plot.InsideTop = 0
            return
        except Exception:                                     # noqa: BLE001
            if attempt == 3:
                raise


def _fit_panel_plot(obj, left: float, top: float, width: float,
                    height: float) -> None:
    """Grow the chart until its *plot* is the size the rows are.

    Asking for a plot area the full size of the chart does not give you one:
    Excel keeps a margin and quietly returns 289pt inside a 300pt chart, which
    divides twenty rows into bands of 14.45pt against a row of 15. Over twenty
    rows that is most of a row of drift, and the bars end up reading against
    the wrong country.

    So the size is solved rather than asserted - enlarge the chart by whatever
    the plot came up short, move it by whatever inset Excel kept, and measure
    again. Three passes settle it. This is the same lesson as ``lock_plot_area``
    passing ``plot_extent`` in rather than deriving it from ChartArea.Height:
    what Excel gives back is not what you asked for, so ask, look, and adjust.
    """
    for _ in range(4):
        _lock_panel_plot(chart := obj.Chart, obj.Width, obj.Height)
        plot = chart.PlotArea
        dh, dw = height - plot.InsideHeight, width - plot.InsideWidth
        dt, dl = plot.InsideTop, plot.InsideLeft
        if abs(dh) < 0.25 and abs(dw) < 0.25 and abs(dt) < 0.25 and abs(dl) < 0.25:
            return
        obj.Height = obj.Height + dh
        obj.Width = obj.Width + dw
        obj.Top = obj.Top - dt
        obj.Left = obj.Left - dl
    _lock_panel_plot(obj.Chart, obj.Width, obj.Height)


def verify_panels(sheet, t: D.Template, layout: L.TableLayout,
                  objects: list) -> list[str]:
    """Prove each panel covers exactly the rows it claims to describe.

    A panel that is one row tall too many draws every bar slightly off its row,
    and the error is invisible until someone reads a number against the wrong
    country. So the check is arithmetic on what Excel actually laid out: the
    category band has to equal the row height, to within a rounding.
    """
    problems: list[str] = []
    rows = len(t.rows)
    tops = {f"panel_{e.tier.key}": sheet.Cells(layout.first_row, n).Top
            for n, e in _panel_columns_of(t, layout)}
    for obj in objects:
        plot = obj.Chart.PlotArea
        band = plot.InsideHeight / rows
        if abs(band - layout.row_height) > 0.15:
            problems.append(
                f"{layout.template_id}: {obj.Name} divides {plot.InsideHeight:.1f}pt "
                f"of plot into {rows} bands of {band:.2f}pt, but a row is "
                f"{layout.row_height:.2f}pt - the bars will not sit on their rows")
        # And the band has to start where the rows do, not merely match in size.
        drift = obj.Top + plot.InsideTop - tops[obj.Name]
        if abs(drift) > 0.5:
            problems.append(
                f"{layout.template_id}: {obj.Name} starts {drift:+.1f}pt from "
                f"the first row it describes")
    return problems


# --------------------------------------------------------------------------- #
# Structure sheets
# --------------------------------------------------------------------------- #
#
# A third kind of sheet, and the simplest of them: one stacked column chart per
# panel, fed from a block of typed segment values with the totals on formulas.
#
# What it has to get right is the thing the template is for - the three panels
# share a scale. Excel will happily give each chart its own, so the maximum is
# computed once from the largest total on the sheet and set on all three. That
# is the one number a reader carries between panels, and a chart library's
# default is to destroy it.


XL_COLUMN_STACKED_100 = 53


def _structure_blocks(t: D.Template, layout: L.StructureLayout) -> dict:
    """Where each panel's data block starts, top to bottom, one under the other."""
    blocks, row = {}, layout.first_row
    for panel in t.structure_panels:
        blocks[panel.key] = row
        row += len(panel.segments) + 3      # header, segments, total, a spacer
    return blocks


def _structure_subtotal(t: D.Template, row: int, index: int) -> str:
    """The SUM behind a subtotal column, mirroring waterfall_spans.

    Same rule as everywhere else in this project, transposed: a subtotal that
    spans n categories sums those n, and one that spans "zero" sums every
    element category before it - skipping the other subtotals, so nothing is
    counted twice.
    """
    spec = t.rows[index]
    if spec.spans == "zero":
        wanted = [i for i in range(index) if t.rows[i].kind != "subtotal"]
    else:
        wanted, seen = [], 0
        for i in range(index - 1, -1, -1):
            if seen >= int(spec.spans):
                break
            if t.rows[i].kind == "subtotal":
                continue
            wanted.append(i)
            seen += 1
        wanted.reverse()

    parts, run = [], [wanted[0]]
    for i in wanted[1:]:
        if i == run[-1] + 1:
            run.append(i)
        else:
            parts.append(run); run = [i]
    parts.append(run)
    return "=" + "+".join(
        f"{L.col_letter(2 + p[0])}{row}" if len(p) == 1
        else f"SUM({L.col_letter(2 + p[0])}{row}:{L.col_letter(2 + p[-1])}{row})"
        for p in parts)


def _write_structure_data(sheet, t: D.Template, layout: L.StructureLayout) -> dict:
    """The typed segment values, and a total under each column that is a formula."""
    blocks = _structure_blocks(t, layout)
    sheet.Columns(1).ColumnWidth = layout.label_width
    for i in range(2, 2 + layout.max_categories):
        sheet.Columns(i).ColumnWidth = layout.value_width

    for panel in t.structure_panels:
        top = blocks[panel.key]
        head = sheet.Cells(top, 1)
        head.Value = panel.label or f"{t.id}{t.variant}"
        head.Font.Bold = True

        for j, (name, scenario) in enumerate(zip(panel.categories,
                                                 panel.category_scenarios)):
            cell = sheet.Cells(top, 2 + j)
            cell.NumberFormat = "@"
            cell.Value = name if layout.horizontal else f"{name} {scenario}"
            cell.HorizontalAlignment = XL_RIGHT_ALIGN

        for i, segment in enumerate(panel.segments):
            r = top + 1 + i
            label = sheet.Cells(r, 1)
            label.NumberFormat = "@"
            label.Value = segment.label
            for j in range(len(panel.categories)):
                cell = sheet.Cells(r, 2 + j)
                cell.NumberFormat = layout.number_format
                row_kind = (t.rows[j].kind if j < len(t.rows) else "element")
                if row_kind == "subtotal":
                    # A total with components is a formula. On this sheet the
                    # subtotals are *columns*, and they are the two rows the
                    # chart does not draw - so a typed one would look right on
                    # the page and be stale in the only place it is read.
                    cell.Formula = _structure_subtotal(t, r, j)
                    cell.Font.Bold = True
                else:
                    cell.Value = segment.values[j]
                    cell.Interior.Color = rgb("#FFF2CC")

        total_row = top + 1 + len(panel.segments)
        total = sheet.Cells(total_row, 1)
        total.NumberFormat = "@"
        total.Value = "Total"
        total.Font.Bold = True
        for j in range(len(panel.categories)):
            cell = sheet.Cells(total_row, 2 + j)
            cell.NumberFormat = layout.number_format
            cell.Font.Bold = True
            # A total with components is a formula. Same rule as everywhere
            # else, and here it is also what the label over the column reads.
            cell.Formula = (f"=SUM({L.col_letter(2 + j)}{top + 1}:"
                            f"{L.col_letter(2 + j)}{total_row - 1})")
    return blocks


def _write_structure_scale(sheet, t: D.Template, layout: L.StructureLayout,
                           blocks: dict, maximum: float) -> tuple[int, str]:
    """Mirror every drawn row as a scaled copy and a text copy, below the data.

    The structure sheets were already fitted - `maximum` is computed from the
    figures rather than declared - but fitted at *build* time, from the Python
    template rather than from the cells. Retyping a segment moved the bar and
    left the axis where it was.

    Below rather than beside. The charts are placed from the width of columns
    A to the last category, so anything written to the right of the data would
    sit underneath them; and rows can be hidden without moving a chart, which
    columns here cannot.

    Returns the row offset and the span cell's address. A scaled copy of row
    `r` lives at `r + shift` and its text copy at `r + 2 * shift`, so the mirror
    keeps the column layout of the original - which is what lets the same
    multi-area range helper read it.
    """
    bottom = max(top + 1 + len(panel.segments)
                 for panel, top in ((p, blocks[p.key])
                                    for p in t.structure_panels))
    shift = bottom + 4
    n = layout.max_categories

    # From the columns that are actually drawn. A horizontal panel's subtotal
    # columns - Europe, Americas, World - are the integrated legend and no bar
    # carries them, so a span taken over them would be set by a number nothing
    # on the chart draws: C02A's World total is 3,733 where the tallest country
    # is 496. Same exclusion the build-time fit already made.
    totals = ",".join(
        _element_range(sheet, t,
                       blocks[panel.key] + 1 + len(panel.segments), 2,
                       len(panel.categories))
        if layout.horizontal else
        f"{L.col_letter(2)}{blocks[panel.key] + 1 + len(panel.segments)}:"
        f"{L.col_letter(1 + len(panel.categories))}"
        f"{blocks[panel.key] + 1 + len(panel.segments)}"
        for panel in t.structure_panels)
    span_cell = sheet.Cells(bottom + 2, 2)
    # The tallest column any panel draws. AGGREGATE ignores the errors a blank
    # category leaves behind; the guard keeps a sheet of nothing from dividing
    # by zero.
    span_cell.Formula = f"=MAX(AGGREGATE(4,6,{totals}),1E-9)"
    span_cell.NumberFormat = "0.000000"
    sheet.Cells(bottom + 2, 1).Value = "span"
    span_ref = f"${L.col_letter(2)}${bottom + 2}"

    # Below which share of the axis a band is too thin to hold its number. The
    # test was made in Python against the build-time figures, which froze it;
    # as a formula it follows an edit. `maximum` is the axis top in data units,
    # so a band's share of the axis is simply its value over it.
    floor = layout.min_label_fraction * maximum

    for panel in t.structure_panels:
        top = blocks[panel.key]
        rows = list(range(top + 1, top + 2 + len(panel.segments)))
        for r in rows:
            for j in range(len(panel.categories)):
                col = 2 + j
                raw = f"{L.col_letter(col)}{r}"
                sheet.Cells(r + shift, col).Formula = (
                    f"=IF(ISBLANK({raw}),NA(),{raw}/{span_ref})")
                sheet.Cells(r + 2 * shift, col).Formula = (
                    f'=IF(ISBLANK({raw}),"",IF({raw}<{floor:.10g},"",'
                    f'TEXT({raw},"{layout.number_format}")))')
    for r in range(shift + 1, 2 * shift + bottom + 2):
        sheet.Rows(r).Hidden = True
    return shift, span_ref


def _element_range(sheet, t: D.Template, row: int, first_col: int,
                   count: int) -> str:
    """An address covering the element rows only, skipping the subtotals.

    A horizontal structure chart plots countries, not continents: the subtotal
    rows are the integrated legend and a bar for Europe would be six times the
    longest country and take the scale with it. Excel accepts a multi-area
    address for a series, so they are simply left out rather than moved.
    """
    wanted = [first_col + i for i, r in enumerate(t.rows[:count])
              if r.kind != "subtotal"]
    parts, run = [], [wanted[0]]
    for c in wanted[1:]:
        if c == run[-1] + 1:
            run.append(c)
        else:
            parts.append(run); run = [c]
    parts.append(run)
    return ",".join(f"${L.col_letter(p[0])}${row}:${L.col_letter(p[-1])}${row}"
                    for p in parts)


def _add_structure_chart(sheet, t: D.Template, layout: L.StructureLayout,
                         panel: D.StructurePanel, top: int, left: float,
                         maximum: float, shift: int = 0, span: float = 1.0):
    """One panel as a stacked column chart, on the scale every panel shares."""
    total_row = top + 1 + len(panel.segments)
    width = layout.panel_width * len(panel.categories) / layout.max_categories
    obj = sheet.ChartObjects().Add(left, layout.chart_top,
                                   max(width, layout.min_panel_width),
                                   layout.chart_height)
    obj.Name = f"panel_{panel.key}"
    chart = obj.Chart
    chart.ChartType = XL_BAR_STACKED if layout.horizontal else XL_COLUMN_STACKED
    # The scale mirror is on hidden rows, and Excel will not plot a hidden cell:
    # a series reading one comes back with no points at all, and the first thing
    # to touch Points(i) fails with "Parameter not valid" rather than anything
    # about visibility. Set before the series are added.
    chart.PlotVisibleOnly = False

    if layout.horizontal:
        # The subtotal rows are the integrated legend, not bars.
        cats = sheet.Range(_element_range(sheet, t, top, 2,
                                          len(panel.categories)))
    else:
        cats = sheet.Range(sheet.Cells(top, 2),
                           sheet.Cells(top, 1 + len(panel.categories)))
    for i, segment in enumerate(panel.segments):
        r = top + 1 + i
        series = chart.SeriesCollection().NewSeries()
        series.Name = segment.label
        drawn = r + shift
        series.Values = (
            sheet.Range(_element_range(sheet, t, drawn, 2,
                                       len(panel.categories)))
            if layout.horizontal else
            sheet.Range(sheet.Cells(drawn, 2),
                        sheet.Cells(drawn, 1 + len(panel.categories))))
        series.XValues = cats
        colour = S.structure_colour(i, len(panel.segments))
        series.Format.Fill.Visible = MSO_TRUE
        series.Format.Fill.Solid()
        series.Format.Fill.ForeColor.RGB = rgb(colour)
        series.Format.Line.Visible = MSO_FALSE
        series.HasDataLabels = True
        labels = series.DataLabels()
        labels.NumberFormat = layout.number_format
        labels.Font.Size = 9
        labels.Font.Name = "Arial"
        labels.Font.Color = rgb(S.on_fill(colour))
        labels.Position = XL_LABEL_CENTER

        # A label only where the band can hold one - suppressing it is a
        # required feature of this template rather than an omission. The test is
        # the band's share of the axis, which is the same test the SVG makes in
        # pixels: 0.9 and 2.9 of 120 go, 3.9 stays, exactly as the reference.
        plotted = [i for i, row in enumerate(t.rows[:len(panel.categories)])
                   if row.kind != "subtotal"] if layout.horizontal else                   list(range(len(panel.categories)))
        for j, source in enumerate(plotted, start=1):
            point = series.Points(j)
            # Linked to the text mirror, which decides for itself whether the
            # band is wide enough to hold a number - the test used to be made
            # here against the build-time figures, which froze it. An empty
            # string is how the cell says "too thin"; the label object stays,
            # so it comes back when the band grows.
            try:
                point.DataLabel.Formula = (
                    f"='{sheet.Name}'!"
                    f"{a1_abs(r + 2 * shift, 2 + source)}")
                point.DataLabel.Format.TextFrame2.WordWrap = MSO_FALSE
            except Exception:                                 # noqa: BLE001
                pass

        # The plan column is outlined, and its base band takes the plan fill -
        # the darkest band is what makes a column read as solid, and a solid
        # column reads as measured. Per point, because which column is planned
        # is structure rather than data and must not move when a figure does.
        for j, scenario in enumerate(panel.category_scenarios, start=1):
            if scenario in ("PL", "BU"):
                point = series.Points(j)
                spec = S.scenario_fill(scenario)
                if i == 0:
                    point.Format.Fill.ForeColor.RGB = rgb(spec.fill)
                    point.DataLabel.Font.Color = rgb(S.on_fill(spec.fill))
                point.Format.Line.Visible = MSO_TRUE
                point.Format.Line.ForeColor.RGB = rgb(spec.outline)
                point.Format.Line.Weight = 0.75

    # The period sum over each column, which the template requires. Carried by
    # a *line* series rather than another band of the stack: a stacked series
    # cannot put a label outside itself - Excel offers only Center, InsideEnd
    # and InsideBase - so a total drawn that way lands on top of the band below
    # it. A line series is exempt, its points sit exactly at the column tops
    # because that is what the values are, and its labels go Above. Nothing is
    # positioned by hand and nothing is written once: the whole thing follows
    # the Total row.
    totals = chart.SeriesCollection().NewSeries()
    totals.Name = "Total"
    totals.Values = sheet.Range(
        sheet.Cells(total_row + shift, 2),
        sheet.Cells(total_row + shift, 1 + len(panel.categories)))
    totals.XValues = cats
    if layout.horizontal:
        # A bar chart refuses a line series - the same limit that costs the pin
        # its head marker - so the total rides a zero-width band stacked at the
        # end. Centre on a band of no width is the bar's end, which is where the
        # figure belongs, and it stays there when the bar moves.
        totals.Values = [0] * len(plotted)
        totals.Format.Fill.Visible = MSO_FALSE
        totals.Format.Line.Visible = MSO_FALSE
        totals.HasDataLabels = True
        marks = totals.DataLabels()
        marks.Font.Size = 9
        marks.Font.Name = "Arial"
        marks.Position = XL_LABEL_CENTER
        for j, source in enumerate(plotted, start=1):
            label = totals.Points(j).DataLabel
            label.Formula = (f"='{sheet.Name}'!"
                             f"${L.col_letter(2 + source)}${total_row}")
            label.Format.Fill.Visible = MSO_TRUE
            label.Format.Fill.Solid()
            label.Format.Fill.ForeColor.RGB = rgb(S.PAGE["background"])
            try:
                label.Format.TextFrame2.WordWrap = MSO_FALSE
            except Exception:                                 # noqa: BLE001
                pass
        chart.ChartGroups(1).GapWidth = layout.gap_width
        strip_chrome(chart, show_categories=True)
        chart.Axes(XL_CATEGORY).ReversePlotOrder = True
        value_axis = chart.Axes(XL_VALUE)
        value_axis.MinimumScale = 0
        value_axis.MaximumScale = maximum / span
        chart.ChartArea.Format.Fill.Visible = MSO_TRUE
        chart.ChartArea.Format.Fill.Solid()
        chart.ChartArea.Format.Fill.ForeColor.RGB = rgb(S.PAGE["background"])
        chart.ChartArea.Format.Line.Visible = MSO_FALSE
        return obj

    totals.ChartType = XL_LINE
    totals.Format.Line.Visible = MSO_FALSE
    totals.MarkerStyle = XL_NONE
    totals.HasDataLabels = True
    marks = totals.DataLabels()
    marks.NumberFormat = layout.number_format
    marks.Font.Size = 9
    marks.Font.Name = "Arial"
    marks.Position = XL_LABEL_ABOVE
    for j in range(1, len(panel.categories) + 1):
        try:
            # The column total, from the text mirror. The series itself plots a
            # fraction of the span, so its own value would print "1.0" over a
            # column of 114.1 - which is the number the whole panel is about.
            totals.Points(j).DataLabel.Formula = (
                f"='{sheet.Name}'!{a1_abs(total_row + 2 * shift, 1 + j)}")
            totals.Points(j).DataLabel.Format.TextFrame2.WordWrap = MSO_FALSE
        except Exception:                                     # noqa: BLE001
            pass

    chart.ChartGroups(1).GapWidth = layout.gap_width
    strip_chrome(chart, show_categories=True)
    # After strip_chrome, which switches the title off - setting it before and
    # re-enabling after gets you Excel's "Chart Title" placeholder back.
    chart.HasTitle = True
    chart.ChartTitle.Text = panel.label
    chart.ChartTitle.Font.Size = 10
    chart.ChartTitle.Font.Name = "Arial"
    chart.ChartTitle.Font.Bold = False
    value_axis = chart.Axes(XL_VALUE)
    value_axis.MinimumScale = 0
    value_axis.MaximumScale = maximum / span   # the shared scale
    # Opaque white. A transparent chart area is fine on the sheet and exports
    # as black, which makes every check of the built workbook a lie.
    chart.ChartArea.Format.Fill.Visible = MSO_TRUE
    chart.ChartArea.Format.Fill.Solid()
    chart.ChartArea.Format.Fill.ForeColor.RGB = rgb(S.PAGE["background"])
    chart.ChartArea.Format.Line.Visible = MSO_FALSE
    return obj


def verify_structure(t: D.Template, objects: list, maximum: float) -> list[str]:
    """Prove the panels really do share a scale.

    The template's whole argument is that 101.9 is the same height in all three
    panels, so this is not a tidiness check: a chart left on its own maximum
    draws its tallest column full height whatever that column is worth, and the
    reader carries a total across and reads it wrong.
    """
    problems: list[str] = []
    scales = {obj.Name: obj.Chart.Axes(XL_VALUE).MaximumScale for obj in objects}
    if len({round(v, 6) for v in scales.values()}) > 1:
        problems.append(f"the panels are on different scales: {scales} - a total "
                        f"carried from one to another would be read wrong")
    for name, value in scales.items():
        if abs(value - maximum) > 1e-6:
            problems.append(f"{name} tops out at {value}, not the {maximum} the "
                            f"largest column on the sheet needs")
    return problems


def build_structure_sheet(excel, sheet, template: D.Template) -> tuple[list[str], list]:
    """One structure template onto one sheet: a data block per panel, then charts."""
    layout = L.structure_layout_for(f"{template.id}{template.variant}")
    sheet.Name = layout.template_id

    _write_input_block(sheet, template, layout, "PL")
    blocks = _write_structure_data(sheet, template, layout)
    excel.Calculate()

    # One scale for every panel, taken from the tallest column on the sheet and
    # rounded up so the axis lands somewhere a reader would choose.
    # From the columns that are actually drawn: a subtotal row left in would
    # set the scale by a number no bar on the chart carries.
    tallest = max(panel.total(i)
                  for panel in template.structure_panels
                  for i in range(len(panel.categories))
                  if not (layout.horizontal and i < len(template.rows)
                          and template.rows[i].kind == "subtotal"))
    step = layout.scale_step
    maximum = math.ceil(tallest / step) * step

    # Derived, not declared - the same rule the chart sheets follow.
    left = sum(sheet.Columns(i).Width
               for i in range(1, 2 + layout.max_categories)) + layout.chart_gap
    # Room for the title block above the panels, which is why the layout is
    # shifted rather than the charts being placed by hand: everything inside
    # _add_structure_chart positions from chart_top.
    title_top, first_left = layout.chart_top, left
    layout = dc.replace(layout, chart_top=layout.chart_top + TITLE_BLOCK_HEIGHT)

    # The scale mirror, and the span it divides by. Read back from the cell
    # rather than reused from `tallest`: the sheet is what the chart reads, and
    # if the two ever disagreed it is the sheet that would be right.
    shift, span_ref = _write_structure_scale(sheet, template, layout, blocks,
                                             maximum)
    span = sheet.Range(span_ref).Value
    if not span:
        raise ValueError(f"{layout.template_id}: the span cell is empty, so "
                         f"every panel would divide by nothing")

    objects = []
    for panel in template.structure_panels:
        obj = _add_structure_chart(sheet, template, layout, panel,
                                   blocks[panel.key], left, maximum,
                                   shift, span)
        objects.append(obj)
        left += obj.Width + layout.panel_gap

    add_title_block(sheet, template, layout, first_left, title_top,
                    max(left - layout.panel_gap - first_left, 260.0))
    problems = verify_structure(template, objects, maximum / span)
    print(f"\n{layout.template_id}: {len(template.structure_panels)} panels "
          f"sharing a scale to {maximum:,.0f}, tallest column {tallest:,.1f}")
    sheet.Range("A1").Select()
    return problems, objects


# --------------------------------------------------------------------------- #
# Line sheets
# --------------------------------------------------------------------------- #
#
# A combo: the monthly tier as columns and the cumulatives as lines, on one
# value axis because they share a unit. Excel will mix column and line series in
# a single chart quite happily - it is only *bar* and line it refuses, which is
# the limit that costs C04A its pin heads and C02A its outside labels.
#
# The cumulative rows are formulas. That is the whole point of the sheet: a line
# chart of running sums where the sums are typed would look right and stop being
# true the moment anyone touched a month.


LINE_SERIES = (
    ("PL month", "monthly", "PL", "column"),
    ("AC month", "monthly", "AC", "column"),
    ("FC month", "monthly", "FC", "column"),
    ("PL cumulative", "cum_pl", "PL", "line"),
    ("AC cumulative", "cum_ac", "AC", "line"),
    ("FC cumulative", "cum_fc", "FC", "line"),
    ("MAT", "mat", None, "line"),
)


def _write_line_data(sheet, t: D.Template, layout: L.LineLayout) -> dict:
    """One row per series; the cumulatives as running sums of the tier above."""
    sheet.Columns(1).ColumnWidth = layout.label_width
    for i in range(2, 14):
        sheet.Columns(i).ColumnWidth = layout.value_width

    head = layout.first_row
    for j, name in enumerate(t.categories):
        cell = sheet.Cells(head, 2 + j)
        cell.NumberFormat = "@"
        cell.Value = name
        cell.HorizontalAlignment = XL_RIGHT_ALIGN
        scen = sheet.Cells(head + 1, 2 + j)
        scen.NumberFormat = "@"
        scen.Value = t.category_scenarios[j]
        scen.Font.Color = rgb(S.PAGE["footnote"])

    rows = {}
    r = head + 2
    for name, key, scenario, _kind in LINE_SERIES:
        rows[name] = r
        label = sheet.Cells(r, 1)
        label.NumberFormat = "@"
        label.Value = name
        tier = t.tier(key)
        series = (tier.series_for(scenario) if scenario
                  else None)
        for j in range(len(t.categories)):
            cell = sheet.Cells(r, 2 + j)
            cell.NumberFormat = layout.number_format
            if key == "monthly":
                value = series.values[j]
                if value is None:
                    continue
                cell.Value = value
                cell.Interior.Color = rgb("#FFF2CC")
            elif key == "mat":
                # The one series that cannot be derived from this page - it
                # reaches back into 2024 - so it is typed like a month.
                value = next((s.values[j] for s in tier.series
                              if s.values[j] is not None), None)
                if value is None:
                    continue
                cell.Value = value
                cell.Interior.Color = rgb("#FFF2CC")
            else:
                source = {"cum_pl": "PL month", "cum_ac": "AC month",
                          "cum_fc": "FC month"}[key]
                if tier.series[0].values[j] is None:
                    # The forecast cumulative carries the handover month too,
                    # so the line joins the actual instead of starting again in
                    # mid-air. It is the same running total either way - the
                    # scenario changes, the series does not.
                    if not (key == "cum_fc"
                            and t.category_scenarios[j] == "AC"
                            and j + 1 < len(t.categories)
                            and t.category_scenarios[j + 1] == "FC"):
                        continue
                # A running sum, and it has to span the actual months as well
                # when it is the forecast line - the forecast cumulates on top
                # of what actually happened, not from zero.
                first = rows[source]
                lo, hi = L.col_letter(2), L.col_letter(2 + j)
                if key == "cum_fc":
                    cell.Formula = (f"=SUM({lo}{rows['AC month']}:{hi}{rows['AC month']})"
                                    f"+SUM({lo}{rows['FC month']}:{hi}{rows['FC month']})")
                else:
                    cell.Formula = f"=SUM({lo}{first}:{hi}{first})"
        r += 1
    return rows


def _write_row_scale(sheet, layout, rows: dict, plan, ncols: int,
                     first_col: int = 2) -> tuple[int, str]:
    """Mirror a row-laid sheet as scaled and text copies, below the data.

    The line sheets put one series per *row* and the periods across columns,
    which is the transpose of everything else in the library - so the scale
    block is a mirrored set of rows rather than a set of columns. The mechanism
    is the same: a span cell, a copy of every plotted row divided by it, and a
    text copy for the labels, because a series plotting a fraction would print
    the fraction and a cell-linked label ignores its own number format.

    `plan` is (row key, the format its label prints in, or None where the
    series carries no label). Returns the row offset and the span cell address;
    a scaled copy of row `r` lives at `r + shift`, its text copy at
    `r + 2 * shift`, so the mirror keeps the column layout of the original.
    """
    # Entries are (key, label format) or (key, label format, scale group).
    # A sheet with one ruler names no group; the driver tree has three - per
    # cent, turnover and currency - and boxes measured in different units must
    # not divide by each other's span.
    entries = [(e[0], e[1], e[2] if len(e) > 2 else "") for e in plan]
    bottom = max(rows.values())
    shift = bottom + 4
    last_col = first_col + ncols - 1

    spans: dict[str, str] = {}
    for i, group in enumerate(dict.fromkeys(g for _k, _f, g in entries)):
        ranges = ",".join(
            f"{L.col_letter(first_col)}{rows[key]}:"
            f"{L.col_letter(last_col)}{rows[key]}"
            for key, _fmt, g in entries if g == group)
        span_row = bottom + 2 + i
        sheet.Cells(span_row, 1).Value = f"span {group}".strip()
        cell = sheet.Cells(span_row, first_col)
        # AGGREGATE ignores the NA()s a scenario mask leaves behind; MAX over an
        # error is an error. Both signs, because a flow that hangs below the
        # axis is as much a length as one that rises above it.
        cell.Formula = (f"=MAX(AGGREGATE(4,6,{ranges}),"
                        f"-AGGREGATE(5,6,{ranges}),1E-9)")
        cell.NumberFormat = "0.000000"
        spans[group] = f"${L.col_letter(first_col)}${span_row}"

    for key, fmt, group in entries:
        span_ref = spans[group]
        r = rows[key]
        sheet.Cells(r + shift, 1).Value = f"{key} scaled"
        for j in range(ncols):
            raw = f"{L.col_letter(first_col + j)}{r}"
            sheet.Cells(r + shift, first_col + j).Formula = (
                f"=IF(ISNA({raw}),NA(),IF(ISBLANK({raw}),NA(),"
                f"{raw}/{span_ref}))")
        if fmt:
            sheet.Cells(r + 2 * shift, 1).Value = f"{key} text"
            escaped = fmt.replace('"', '""')
            for j in range(ncols):
                raw = f"{L.col_letter(first_col + j)}{r}"
                sheet.Cells(r + 2 * shift, first_col + j).Formula = (
                    f'=IF(ISNA({raw}),"",IF(ISBLANK({raw}),"",'
                    f'TEXT({raw},"{escaped}")))')
    for r in range(shift + 1, 2 * shift + bottom + 2):
        sheet.Rows(r).Hidden = True
    return shift, (spans[""] if list(spans) == [""] else spans)


def _add_line_chart(sheet, t: D.Template, layout: L.LineLayout, rows: dict,
                    left: float):
    """The combo chart: three column series and four lines on one axis."""
    head = layout.first_row
    obj = sheet.ChartObjects().Add(left, layout.chart_top, layout.chart_width,
                                   layout.chart_height)
    obj.Name = "line_chart"
    chart = obj.Chart
    chart.ChartType = XL_COLUMN_CLUSTERED
    cats = sheet.Range(sheet.Cells(head, 2), sheet.Cells(head, 13))

    # Every plotted row divided by one span, so the columns and the four lines
    # stay on the ruler they share - which is the whole reason a monthly bar
    # and a cumulative line can sit in one chart and be read together.
    shift, span_ref = _write_row_scale(
        sheet, layout, rows,
        [(name, None if name == "PL month" else "#,##0")
         for name, _k, _s, _kind in LINE_SERIES],
        len(t.categories))
    span = sheet.Range(span_ref).Value
    if not span:
        raise ValueError(f"{layout.template_id}: the span cell is empty, so "
                         f"every series would divide by nothing")
    # The scale mirror is on hidden rows and Excel will not plot a hidden cell.
    chart.PlotVisibleOnly = False

    for name, key, scenario, kind in LINE_SERIES:
        r = rows[name]
        series = chart.SeriesCollection().NewSeries()
        series.Name = name
        series.Values = sheet.Range(sheet.Cells(r + shift, 2),
                                    sheet.Cells(r + shift, 13))
        series.XValues = cats
        if kind == "line":
            series.ChartType = XL_LINE_MARKERS
            series.Format.Line.Visible = MSO_TRUE
            series.Format.Line.ForeColor.RGB = rgb(S.TEXT["primary"])
            series.Format.Line.Weight = 0.75
            series.MarkerStyle = XL_MARKER_SQUARE
            series.MarkerSize = 5
            spec = S.scenario_fill(scenario or "AC")
            series.MarkerBackgroundColor = rgb(spec.fill)
            series.MarkerForegroundColor = rgb(spec.outline or S.TEXT["primary"])
            if name == "FC cumulative":
                # The handover point belongs to the actual line; the forecast
                # only borrows it to join on, so it draws no marker of its own.
                try:
                    series.Points(8).MarkerStyle = XL_NONE
                except Exception:                             # noqa: BLE001
                    pass
        else:
            apply_scenario(series.Format, scenario)
            series.Format.Line.Visible = (MSO_TRUE if scenario == "PL"
                                          else MSO_FALSE)
            if scenario == "PL":
                series.Format.Line.ForeColor.RGB = rgb(
                    S.scenario_fill("PL").outline)
        # IBCS labels the elements rather than running a value axis, which is
        # why strip_chrome takes the axis away - so a chart with the axis gone
        # and no labels states no numbers at all.
        #
        # But not *every* element. Where two lines run close together a label
        # on each says nothing the eye can read, so the reference prints only
        # some of them, and the data layer records which in the tier's
        # `printed` tuple - a None means IBCS drew the point and left the
        # number off. Honouring it here is what stops the Excel MAT line
        # stacking twelve labels into the space the SVG gives six.
        #
        # The plan column keeps its shape but not its number. It stands behind
        # the actual as the thing being compared against, and at this density
        # its label lands on the actual's - the reference prints one number per
        # month for the same reason.
        series.HasDataLabels = name != "PL month"
        if series.HasDataLabels:
            labels = series.DataLabels()
            labels.NumberFormat = "#,##0"
            labels.Font.Size = 8
            labels.Font.Name = "Arial"
            labels.Position = (XL_LABEL_ABOVE if kind == "line"
                               else XL_LABEL_OUTSIDE_END)
            # The number comes from the text mirror, because the series now
            # plots a fraction of the span.
            for j in range(len(t.categories)):
                try:
                    series.Points(j + 1).DataLabel.Formula = (
                        f"='{sheet.Name}'!{a1_abs(r + 2 * shift, 2 + j)}")
                except Exception:                             # noqa: BLE001
                    pass
            printed = t.tier(key).printed
            if printed is not None:
                for j, value in enumerate(printed, start=1):
                    if value is not None:
                        continue
                    try:
                        series.Points(j).HasDataLabel = False
                    except Exception:                         # noqa: BLE001
                        pass          # no point there at all; nothing to hide

    chart.ChartGroups(1).GapWidth = layout.gap_width
    chart.ChartGroups(1).Overlap = layout.overlap
    strip_chrome(chart, show_categories=True)
    axis = chart.Axes(XL_VALUE)
    axis.MinimumScale = 0
    axis.MaximumScale = layout.maximum / span
    chart.ChartArea.Format.Fill.Visible = MSO_TRUE
    chart.ChartArea.Format.Fill.Solid()
    chart.ChartArea.Format.Fill.ForeColor.RGB = rgb(S.PAGE["background"])
    chart.ChartArea.Format.Line.Visible = MSO_FALSE
    return obj


def verify_line(sheet, t: D.Template, layout: L.LineLayout, rows: dict) -> list[str]:
    """The cumulatives must be formulas, and must equal what the data layer says."""
    problems: list[str] = []
    for name, key in (("PL cumulative", "cum_pl"), ("AC cumulative", "cum_ac"),
                      ("FC cumulative", "cum_fc")):
        tier = t.tier(key)
        for j in range(len(t.categories)):
            expected = tier.series[0].values[j]
            cell = sheet.Cells(rows[name], 2 + j)
            if expected is None:
                continue
            if not str(cell.Formula).startswith("="):
                problems.append(f"{layout.template_id}: {name} {t.categories[j]} "
                                f"is typed where it should be a running sum")
            if cell.Value is None or abs(float(cell.Value) - expected) > 0.005:
                problems.append(f"{layout.template_id}: {name} {t.categories[j]} "
                                f"shows {cell.Value} where the data says {expected}")
    return problems


XL_AREA = 1

# C08H's rows, in sheet order. The level is a stock and the other two are flows,
# which is why the level needs an opening balance the flows do not.
C08H_ROWS = (
    ("Increase", "flow", "column"),
    ("Decrease", "outflow", "column"),
    ("Inventory change", "change", "column"),
    ("Inventory", "level", "area"),
)


def _write_c08h_data(sheet, t: D.Template, layout: L.LineLayout) -> dict:
    """Flows typed, change and level derived - a stock that is typed is a lie."""
    sheet.Columns(1).ColumnWidth = layout.label_width
    for i in range(2, 2 + len(t.categories) + 1):
        sheet.Columns(i).ColumnWidth = layout.value_width

    head = layout.first_row
    for j, name in enumerate(t.categories):
        for row, text in ((head, name.split()[0]), (head + 1, name.split()[1])):
            cell = sheet.Cells(row, 2 + j)
            cell.NumberFormat = "@"
            cell.Value = text
            cell.HorizontalAlignment = XL_RIGHT_ALIGN
        scen = sheet.Cells(head + 2, 2 + j)
        scen.NumberFormat = "@"
        scen.Value = t.category_scenarios[j]
        scen.Font.Color = rgb(S.PAGE["footnote"])

    rows, r = {}, head + 3
    for name, key, _kind in C08H_ROWS:
        rows[name] = r
        label = sheet.Cells(r, 1)
        label.NumberFormat = "@"
        label.Value = name
        for j in range(len(t.categories)):
            cell = sheet.Cells(r, 2 + j)
            cell.NumberFormat = layout.number_format
            letter = L.col_letter(2 + j)
            if key in ("flow", "outflow"):
                value = next((sr.values[j] for sr in t.tier(key).series
                              if sr.values[j] is not None), None)
                cell.Value = value
                cell.Interior.Color = rgb("#FFF2CC")
            elif key == "change":
                cell.Formula = (f"={letter}{rows['Increase']}"
                                f"-{letter}{rows['Decrease']}")
            else:
                # The level accumulates: the opening balance plus every change
                # so far. A running sum rather than a walk, so a corrected
                # quarter moves every level after it at once.
                lo = L.col_letter(2)
                change_row = rows["Inventory change"]
                cell.Formula = (f"=$A${r}+SUM({lo}{change_row}:"
                                f"{letter}{change_row})")
        r += 1

    opening = sheet.Cells(rows["Inventory"], 1)
    opening.NumberFormat = layout.number_format
    opening.Value = D.C08H_OPENING
    opening.Interior.Color = rgb("#FFF2CC")

    # Chart feed. Two things Excel will not do from the rows above:
    #
    # An *area* series carries one fill for the whole series - unlike a column
    # series, whose points can be painted individually - so an area cannot
    # change from actual to plan part way along. It has to be three areas, each
    # holding values only where its own scenario runs, and NA() elsewhere so
    # nothing is drawn there. The runs overlap by one quarter, or the areas
    # would meet across a gap.
    #
    # And a decrease has to be *negative* to hang below the axis. The rows hold
    # magnitudes, because that is what the reference prints.
    level = rows["Inventory"]
    for scenario in ("AC", "FC", "PL"):
        r = max(rows.values()) + 1
        rows[f"level {scenario}"] = r
        cell = sheet.Cells(r, 1)
        cell.NumberFormat = "@"
        cell.Value = f"Inventory {scenario}"
        cell.Font.Color = rgb(S.PAGE["footnote"])
        for j, scen in enumerate(t.category_scenarios):
            previous = t.category_scenarios[j - 1] if j else None
            wanted = scen == scenario or previous == scenario
            letter = L.col_letter(2 + j)
            sheet.Cells(r, 2 + j).Formula = (
                f"={letter}{level}" if wanted else "=NA()")
    r = max(rows.values()) + 1
    rows["Decrease drawn"] = r
    cell = sheet.Cells(r, 1)
    cell.NumberFormat = "@"
    cell.Value = "Decrease (drawn)"
    cell.Font.Color = rgb(S.PAGE["footnote"])
    for j in range(len(t.categories)):
        letter = L.col_letter(2 + j)
        sheet.Cells(r, 2 + j).Formula = f"=-{letter}{rows['Decrease']}"
    return rows


def _add_c08h_charts(sheet, t: D.Template, layout: L.LineLayout, rows: dict,
                     left: float) -> list:
    """Two charts: the change tier, and the level with its movements.

    Two rather than three, because the level and the two flows share a zero -
    the increase columns rise into the stock they add to - so only the change
    tier needs an axis of its own. That is a fact about the template rather than
    a saving: drawing them apart would break the one relationship the chart is
    for.
    """
    head = layout.first_row
    cats = sheet.Range(sheet.Cells(head, 2),
                       sheet.Cells(head, 1 + len(t.categories)))
    objects = []

    plan = (
        # (name, height, top, members, transparent)
        ("change", 150.0, layout.chart_top,
         (("Inventory change", "column", +1),), False),
        # The level and the flows share an axis but cannot share a chart:
        # setting one series to area converts the whole group, and the columns
        # come out as stacked areas. So they are two charts in one place - the
        # level behind, the movements in front on a transparent ground - with
        # the same bounds and the same locked plot, which is what keeps their
        # categories over each other.
        ("level", layout.chart_height, layout.chart_top + 160.0,
         (("level AC", "area", None), ("level FC", "area", None),
          ("level PL", "area", None)), False),
        ("flows", layout.chart_height, layout.chart_top + 160.0,
         (("Increase", "column", +1), ("Decrease drawn", "column", -1)), True),
    )

    # The level and the flows are the two charts that declare bounds, and they
    # declare the same ones - the columns rise into the stock they add to, so
    # they are one ruler drawn twice. Those are the rows that divide by a span.
    # The change tier is left alone: it declares nothing and Excel fits it, so
    # it already follows its data.
    scaled_rows = ("level AC", "level FC", "level PL",
                   "Increase", "Decrease drawn")
    shift, span_ref = _write_row_scale(
        sheet, layout, rows,
        [(key, "0" if key in ("Increase", "Decrease drawn") else None)
         for key in scaled_rows],
        len(t.categories))
    span = sheet.Range(span_ref).Value
    if not span:
        raise ValueError(f"{layout.template_id}: the span cell is empty, so "
                         f"the level and its flows would divide by nothing")

    for name, height, top, members, transparent in plan:
        obj = sheet.ChartObjects().Add(left, top, layout.chart_width, height)
        obj.Name = f"c08_{name}"
        chart = obj.Chart
        chart.ChartType = XL_AREA if members[0][1] == "area" else XL_COLUMN_CLUSTERED
        # The scale mirror is on hidden rows, which Excel will not plot from
        # unless told.
        chart.PlotVisibleOnly = False
        for label, kind, sign in members:
            r = rows[label] + (shift if label in scaled_rows else 0)
            series = chart.SeriesCollection().NewSeries()
            series.Name = label
            series.Values = sheet.Range(sheet.Cells(r, 2),
                                        sheet.Cells(r, 1 + len(t.categories)))
            series.XValues = cats
            series.Format.Fill.Visible = MSO_TRUE
            series.Format.Fill.Solid()
            if kind == "area":
                which = label.split()[-1]
                spec = S.scenario_fill(which)
                if spec.hatch:
                    series.Format.Fill.Patterned(MSO_PATTERN_WIDE_UPWARD)
                    series.Format.Fill.ForeColor.RGB = rgb(spec.hatch.colour)
                    series.Format.Fill.BackColor.RGB = rgb(spec.fill)
                else:
                    series.Format.Fill.ForeColor.RGB = rgb(
                        S.STRUCTURE_RAMP[-1] if which == "AC" else "#F2F2F2")
                series.Format.Line.Visible = MSO_FALSE
            series.HasDataLabels = (kind != "area")
            # Per point, because the scenario changes part way along and the
            # fill is what says so.
            for j, scenario in enumerate(t.category_scenarios, start=1):
                point = series.Points(j)
                if kind == "area":
                    continue      # one fill per series; set below
                else:
                    apply_scenario(point.Format, scenario)
                    if not S.scenario_fill(scenario).hatch:
                        point.Format.Fill.ForeColor.RGB = rgb(
                            S.waterfall_fill(scenario, -1 if sign > 0 else +1))
            if series.HasDataLabels:
                marks = series.DataLabels()
                if label in scaled_rows:
                    # The series plots a fraction of the span, so the number
                    # comes from the text mirror instead of from the series.
                    for j in range(len(t.categories)):
                        try:
                            series.Points(j + 1).DataLabel.Formula = (
                                f"='{sheet.Name}'!"
                                f"{a1_abs(rows[label] + 2 * shift, 2 + j)}")
                        except Exception:                     # noqa: BLE001
                            pass
                for attempt in (
                        lambda: setattr(marks, "NumberFormat",
                                        "+0;-0;0" if label == "Inventory change" else "0"),
                        lambda: setattr(marks.Font, "Size", 8),
                        lambda: setattr(marks.Font, "Name", "Arial"),
                        lambda: setattr(marks, "Position", XL_LABEL_OUTSIDE_END)):
                    try:
                        attempt()
                    except Exception:                         # noqa: BLE001
                        pass
        try:
            chart.ChartGroups(1).GapWidth = layout.gap_width
            if name == "flows":
                # One slot, not two: an increase and a decrease belong to the
                # same quarter and sit above and below its axis, not beside
                # each other.
                chart.ChartGroups(1).Overlap = 100
        except Exception:                                     # noqa: BLE001
            pass
        strip_chrome(chart, show_categories=(name == "flows"))
        if name in ("level", "flows"):
            axis = chart.Axes(XL_VALUE)
            axis.MinimumScale = -8 / span
            axis.MaximumScale = layout.maximum / span
            _lock_panel_plot(chart, obj.Width - 8, obj.Height - 30)
        chart.ChartArea.Format.Line.Visible = MSO_FALSE
        if transparent:
            chart.ChartArea.Format.Fill.Visible = MSO_FALSE
            chart.PlotArea.Format.Fill.Visible = MSO_FALSE
        else:
            chart.ChartArea.Format.Fill.Visible = MSO_TRUE
            chart.ChartArea.Format.Fill.Solid()
            chart.ChartArea.Format.Fill.ForeColor.RGB = rgb(S.PAGE["background"])
        objects.append(obj)
    return objects


def verify_c08h(sheet, t: D.Template, layout: L.LineLayout, rows: dict) -> list[str]:
    """The level must be a formula, and must come back to the printed figures.

    This is the check the template earns: a stock is the one quantity on the
    page that nobody measures directly, so if it is typed it will be right on
    the day it is written and wrong ever after.
    """
    problems: list[str] = []
    levels = (D.C08H_OPENING,) + D.C08H_LEVELS
    for j in range(len(t.categories)):
        cell = sheet.Cells(rows["Inventory"], 2 + j)
        if not str(cell.Formula).startswith("="):
            problems.append(f"{layout.template_id}: the level at "
                            f"{t.categories[j]} is typed, not accumulated")
        want = levels[j + 1]
        if cell.Value is None or abs(float(cell.Value) - want) > 0.005:
            problems.append(f"{layout.template_id}: the level at "
                            f"{t.categories[j]} shows {cell.Value}, not {want}")
    return problems


def build_line_sheet(excel, sheet, template: D.Template) -> tuple[list[str], list]:
    """One line template onto one sheet: the series as rows, then a combo chart."""
    layout = L.line_layout_for(f"{template.id}{template.variant}")
    sheet.Name = layout.template_id

    _write_input_block(sheet, template, layout, "PL")
    # Room for the title block, taken off the top of the chart zone. Both line
    # sheets position everything from chart_top, so shifting the layout moves
    # the charts and nothing else has to know.
    title_top = layout.chart_top
    layout = dc.replace(layout, chart_top=layout.chart_top + TITLE_BLOCK_HEIGHT)
    if template.id == "C08":
        rows = _write_c08h_data(sheet, template, layout)
        excel.Calculate()
        left = (sum(sheet.Columns(i).Width
                    for i in range(1, 2 + len(template.categories)))
                + layout.chart_gap)
        objects = _add_c08h_charts(sheet, template, layout, rows, left)
        add_title_block(sheet, template, layout, left, title_top,
                        layout.chart_width)
        excel.Calculate()
        problems = verify_c08h(sheet, template, layout, rows)
        print(chr(10) + f"{layout.template_id}: two charts - the change "
              f"tier, and the level with the movements it shares a zero with")
        sheet.Range("A1").Select()
        return problems, objects
    rows = _write_line_data(sheet, template, layout)
    excel.Calculate()
    left = sum(sheet.Columns(i).Width for i in range(1, 14)) + layout.chart_gap
    obj = _add_line_chart(sheet, template, layout, rows, left)
    add_title_block(sheet, template, layout, left, title_top,
                    layout.chart_width)
    excel.Calculate()

    problems = verify_line(sheet, template, layout, rows)
    print(f"\n{layout.template_id}: {len(LINE_SERIES)} series - three column, "
          f"four line - on one axis to {layout.maximum:,.0f}")
    sheet.Range("A1").Select()
    return problems, [obj]


def build_table_sheet(excel, sheet, template: D.Template) -> tuple[list[str], list]:
    """One table onto one sheet. No charts, and therefore no chart zone."""
    layout = L.table_layout_for(f"{template.id}{template.variant}")
    sheet.Name = layout.template_id

    plan = layout.plan(template)
    for i, entry in enumerate(plan, start=1):
        if entry.kind == "label":
            width = layout.label_width
        elif entry.tier.key in layout.panels:
            # As many times wider than a printed column as the reference draws
            # it, so the panels keep the scale they share.
            width = layout.panel_value_width * layout.panels[entry.tier.key].width_px / 64.0
        else:
            width = layout.value_width
        sheet.Columns(i).ColumnWidth = width

    _write_table_title(sheet, template, layout)
    _write_table_header(sheet, template, layout)
    _write_table_body(sheet, template, layout)
    thresholds = _write_table_footnote(sheet, template, layout)
    excel.Calculate()
    _apply_table_red(sheet, template, layout, thresholds)
    excel.Calculate()

    objects = []
    if template.panel_tiers:
        scaffolding = _write_panel_scaffolding(sheet, template, layout)
        excel.Calculate()
        objects = _add_panel_charts(sheet, template, layout, scaffolding)

    problems = verify_table(sheet, template, layout)
    problems += verify_panels(sheet, template, layout, objects)
    print(f"\n{layout.template_id}: {len(template.rows)} rows x "
          f"{len(plan) - 1} columns, rows {layout.first_row}-"
          f"{layout.last_row(template)}, thresholds on row "
          f"{layout.footnote_row(template)}")
    sheet.Range("A1").Select()
    return problems, objects


# --------------------------------------------------------------------------- #
# XY sheets
# --------------------------------------------------------------------------- #
#
# The one sheet in the project with no formula on it, and it is worth saying why
# rather than treating it as an oversight. Every other template here computes
# something - a variance, a subtotal, a running sum - and the standing rule is
# that a computed number must be a formula. A portfolio chart computes nothing:
# market attractiveness is a score somebody assigned, relative market share is a
# ratio against a competitor whose sales are not on this page, and net sales is
# money. Three independent measurements, no arithmetic between them.
#
# So "every total is a formula" has nothing to bite on, and the claim the sheet
# has to prove instead is the other half of live: that the *chart* reads the
# cells. A chart whose series were pasted as constants - which is what happens
# when a chart is copied between workbooks - looks identical and stops following
# its data, and on a sheet with no formulas that is the only way to be wrong.

XL_BUBBLE = 15
XL_SIZE_IS_AREA = 1
XL_SIZE_IS_WIDTH = 2      # never set; named so verify_xy's check reads as a choice
XL_LABEL_CENTRE = -4108


def _write_xy_inputs(sheet, t: D.Template, layout: L.XYLayout) -> None:
    """The title inputs, plus the two that name the third measure.

    A bubble chart's size channel needs a measure and a unit of its own, and
    they are inputs like any other: the chart's legend heading and the sheet's
    own column header are both formulas over them, so there is one place to
    correct "mEUR" and it corrects both.
    """
    _write_input_block(sheet, t, layout, "PY")
    if t.size_legend is None:
        # A scattergram has no size channel: its third measure is derived from
        # the two axes rather than drawn, so there is nothing to name here.
        return
    measure, unit = t.size_legend
    label = sheet.Cells(layout.size_row, 1)
    label.Value = "Size"
    label.Font.Color = rgb("#808080")
    for column, value in ((2, measure), (3, unit)):
        cell = sheet.Cells(layout.size_row, column)
        cell.NumberFormat = "@"
        cell.Value = value
        cell.Interior.Color = rgb("#FFF2CC")


def _write_xy_data(sheet, t: D.Template, layout: L.XYLayout) -> dict:
    """One row per point, grouped into a block per series.

    Grouped because an Excel series takes one fill, so everything that shares a
    fill has to be a contiguous range before it can be a series - the same
    constraint the stacked charts run into, arriving from the other direction.

    Five columns, the same five for both templates: the point's name, its two
    coordinates, what makes it a series, and its third measure. That last pair
    is where the two templates differ and it is worth reading them side by side.
    A bubble chart's series key is the *scenario* and its third measure is a
    number, the bubble's area; a scattergram's series key is a *category* - the
    product line - and its third measure is derived from the other two. Same
    shape, different notation, and the sheet says which without being told.
    """
    sheet.Columns(1).ColumnWidth = layout.label_width
    for i in (2, 3, 4, 5):
        sheet.Columns(i).ColumnWidth = layout.value_width

    axis_x, axis_y = t.axes
    head = layout.first_row
    sized = any(p.size is not None for p in t.points)
    captions = [(1, "Name"), (2, _axis_caption(axis_x)), (3, _axis_caption(axis_y)),
                (4, "Scenario" if layout.series_key == "scenario" else "Group")]
    for column, caption in captions:
        cell = sheet.Cells(head, column)
        cell.NumberFormat = "@"
        cell.Value = caption
        cell.Font.Bold = True
    if sized:
        # The size column heads itself from the two size inputs, so the sheet
        # says what a bubble's area means in the same words the legend does.
        size_head = sheet.Cells(head, 5)
        size_head.Formula = f'=B{layout.size_row}&" in "&C{layout.size_row}'
        size_head.Font.Bold = True

    rows: dict[str, tuple[int, int]] = {}
    for key in layout.scenarios(t):
        first, last = layout.block(t, key)
        rows[key] = (first, last)
        tag = sheet.Cells(first - 1, 1)
        tag.NumberFormat = "@"
        tag.Value = key
        tag.Font.Color = rgb(S.PAGE["footnote"])
        points = [p for p in t.points if layout.key_of(p) == key]
        # A block at a time, not a cell at a time. C09 writes 149 rows, and
        # setting five cells apiece over COM is enough traffic that Excel starts
        # refusing calls outright - "Call was rejected by callee", which reads
        # like a bug in the code rather than the queue backing up.
        def block(column, values, fmt, shade=False):
            rng = sheet.Range(sheet.Cells(first, column), sheet.Cells(last, column))
            rng.NumberFormat = fmt
            rng.Value = tuple((v,) for v in values)
            if shade:
                rng.Interior.Color = rgb("#FFF2CC")

        block(1, [p.entity for p in points], "@")
        block(2, [p.x for p in points], layout.coordinate_format, shade=True)
        block(3, [p.y for p in points], layout.coordinate_format, shade=True)
        block(4, [key] * len(points), "@")
        if sized:
            block(5, [p.size for p in points], layout.size_format, shade=True)
    return rows


def _axis_caption(axis: D.Axis) -> str:
    return f"{axis.title} in {axis.unit}" if axis.unit else axis.title


def _xy_series_colour(t: D.Template, layout: L.XYLayout, key: str) -> str:
    """The fill one series takes, and the whole notational point of the pair.

    A bubble chart's series is a scenario, so it takes a scenario fill and needs
    no legend: solid light is a prior year on every page ever printed. A
    scattergram's series is a product line, which means nothing outside this
    chart, so it takes an accent - a colour that is deliberately not one of the
    scenario greys, and comes with a key.
    """
    if layout.series_key == "scenario":
        return S.bubble_fill(key)[0]
    return S.structure_colour(D.C09C_ACCENT[key], accent=True)


def _add_xy_chart(sheet, t: D.Template, layout: L.XYLayout, rows: dict,
                  left: float, extras: dict | None = None):
    """One chart: a series per key, plus whatever the template draws over it.

    A bubble chart and a scattergram are the same chart with one difference -
    whether the third measure is the size of the mark or a curve behind it - so
    they share everything here except the chart type and that.

    ``SizeRepresents`` is set explicitly on the bubble chart even though area is
    Excel's default. It is the notation: width-proportional bubbles overstate a
    large value by the square, and a default is not a decision anyone can see.
    """
    sized = any(p.size is not None for p in t.points)
    obj = sheet.ChartObjects().Add(left, layout.chart_top,
                                   layout.chart_width, layout.chart_height)
    obj.Name = f"xy_{layout.template_id}"
    chart = obj.Chart
    chart.ChartType = XL_BUBBLE if sized else XL_XY_SCATTER

    axis_x, axis_y = t.axes
    for key in layout.scenarios(t):
        first, last = rows[key]
        series = chart.SeriesCollection().NewSeries()
        series.Name = key
        series.XValues = sheet.Range(sheet.Cells(first, 2), sheet.Cells(last, 2))
        series.Values = sheet.Range(sheet.Cells(first, 3), sheet.Cells(last, 3))
        if sized:
            # BubbleSizes wants an address string; handing it a Range object is
            # rejected, and the message names the Values property rather than
            # this one, which sends you looking in the wrong place.
            series.BubbleSizes = f"='{sheet.Name}'!$E${first}:$E${last}"

        colour = _xy_series_colour(t, layout, key)
        opacity = S.bubble_fill(key)[1] if sized else 0.85
        if sized:
            series.Format.Fill.Visible = MSO_TRUE
            series.Format.Fill.Solid()
            series.Format.Fill.ForeColor.RGB = rgb(colour)
            series.Format.Fill.Transparency = 1.0 - opacity
            series.Format.Line.Visible = MSO_FALSE
        else:
            series.MarkerStyle = XL_MARKER_CIRCLE
            series.MarkerSize = 6
            series.MarkerBackgroundColor = rgb(colour)
            series.MarkerForegroundColor = rgb("#404040")
            series.Format.Line.Visible = MSO_FALSE

        points = [p for p in t.points if layout.key_of(p) == key]
        labelled = [(n, p) for n, p in enumerate(points, start=1)
                    if sized or p.entity]
        series.HasDataLabels = bool(labelled)
        if not labelled:
            continue
        marks = series.DataLabels()
        marks.Font.Size = 8
        marks.Font.Name = "Arial"
        for n, point in enumerate(points, start=1):
            label = series.Points(n).DataLabel
            if sized:
                # Name and value in one label. The page puts the name above the
                # bubble and the value inside it, which needs two labels on one
                # point; Excel allows one, so the workbook stacks them and says
                # the same two things.
                label.Position = XL_LABEL_CENTRE
                label.Font.Color = rgb(S.on_fill(colour))
                label.Text = f"{point.entity}{chr(10)}{point.size:,.1f}"
            elif point.entity:
                label.Position = XL_LABEL_ABOVE
                label.Text = point.entity
            else:
                label.Text = ""
                label.ShowValue = False
                label.ShowSeriesName = False
                label.ShowCategoryName = False

    if extras is not None:
        _add_c09_curves(chart, sheet, extras)

    if sized:
        group = chart.ChartGroups(1)
        group.SizeRepresents = XL_SIZE_IS_AREA
        group.BubbleScale = layout.bubble_scale

    chart.HasLegend = False
    chart.HasTitle = False
    for axis_id, axis in ((XL_CATEGORY, axis_x), (XL_VALUE, axis_y)):
        drawn = chart.Axes(axis_id)
        # Declared, never fitted. Excel left to choose would pick its own round
        # numbers, and a share axis that stops at 1.20 says something different
        # about the same nine units from one that stops at 1.25.
        drawn.MinimumScale = axis.minimum
        drawn.MaximumScale = axis.maximum
        drawn.MajorUnit = axis.step
        drawn.HasMajorGridlines = True
        drawn.MajorGridlines.Format.Line.ForeColor.RGB = rgb("#A6A6A6")
        drawn.MajorTickMark = XL_TICK_MARK_NONE
        drawn.MinorTickMark = XL_TICK_MARK_NONE
        drawn.TickLabels.Font.Size = 9
        drawn.TickLabels.Font.Name = "Arial"
        drawn.TickLabels.NumberFormat = ("0.00" if axis.maximum <= 2 else "0")
        drawn.HasTitle = True
        drawn.AxisTitle.Text = _axis_caption(axis)
        drawn.AxisTitle.Font.Size = 9
        drawn.AxisTitle.Font.Bold = True
        drawn.AxisTitle.Font.Name = "Arial"

    chart.PlotArea.Format.Line.Visible = MSO_FALSE
    chart.PlotArea.Format.Fill.Visible = MSO_FALSE
    chart.ChartArea.Format.Fill.Visible = MSO_TRUE
    chart.ChartArea.Format.Fill.Solid()
    chart.ChartArea.Format.Fill.ForeColor.RGB = rgb(S.PAGE["background"])
    chart.ChartArea.Format.Line.Visible = MSO_FALSE
    return obj


def verify_xy(sheet, t: D.Template, layout: L.XYLayout, rows: dict,
              objects: list) -> list[str]:
    """The chart must read the cells, and both axes must be as declared.

    On a sheet with no formulas these are the only two ways to be wrong, so
    they are checked rather than assumed: a series bound to pasted constants,
    and an axis Excel scaled for itself.
    """
    problems: list[str] = []
    chart = objects[0].Chart
    axis_x, axis_y = t.axes

    for n, scenario in enumerate(layout.scenarios(t), start=1):
        first, last = rows[scenario]
        series = chart.SeriesCollection(n)
        formula = str(series.Formula)
        for what, want in (("x", f"$B${first}:$B${last}"),
                           ("y", f"$C${first}:$C${last}")):
            if want not in formula:
                problems.append(
                    f"{layout.template_id}: the {scenario} series does not read "
                    f"its {what} from {want} - it is {formula}")
        count = sum(1 for p in t.points if layout.key_of(p) == scenario)
        if last - first + 1 != count:
            problems.append(f"{layout.template_id}: {scenario} has {count} "
                            f"points over {last - first + 1} rows")

    if (any(p.size is not None for p in t.points)
            and chart.ChartGroups(1).SizeRepresents != XL_SIZE_IS_AREA):
        problems.append(f"{layout.template_id}: bubble size represents width, "
                        f"not area - every large value is overstated")

    for axis_id, axis, name in ((XL_CATEGORY, axis_x, "x"),
                                (XL_VALUE, axis_y, "y")):
        drawn = chart.Axes(axis_id)
        for what, got, want in (("minimum", drawn.MinimumScale, axis.minimum),
                                ("maximum", drawn.MaximumScale, axis.maximum),
                                ("step", drawn.MajorUnit, axis.step)):
            if abs(float(got) - want) > 1e-9:
                problems.append(f"{layout.template_id}: the {name} axis {what} "
                                f"is {got}, not the declared {want}")
    return problems


XL_XY_SCATTER = -4169
C09C_CURVE_STEPS = 60


def _write_c09_extras(sheet, t: D.Template, layout: L.XYLayout,
                      rows: dict) -> dict:
    """Gross profit per product, the segment count, and the curve series.

    This is where the scattergram earns its formulas. The bubble chart had
    none - three independent measurements and no arithmetic between them - but
    a scattergram whose axes multiply into a third measure has plenty, and every
    one of them is the kind that must never be typed:

      * gross profit is net sales times margin, per product;
      * the count the message states is how many of one product line clear a
        threshold on that product;
      * and each curve is the locus of a constant gross profit, so its net
        sales at a given margin is a division, not a series of typed points.

    Type any of the three and the sheet keeps saying 45 after the data stops
    supporting it, which is precisely the failure the message makes possible.
    """
    head = layout.first_row
    gp_col, first_row, last_row = 5, min(r[0] for r in rows.values()), max(r[1] for r in rows.values())

    caption = sheet.Cells(head, gp_col)
    caption.NumberFormat = "@"
    caption.Value = "Gross profit in mUSD"
    caption.Font.Bold = True
    # Walk the blocks rather than the span: most products have no name, so a
    # test on the label column would skip nearly every row on the sheet.
    for first, last in rows.values():
        rng = sheet.Range(sheet.Cells(first, gp_col), sheet.Cells(last, gp_col))
        rng.NumberFormat = "0.00"
        # One relative formula filled down the block; Excel adjusts the rows.
        rng.Formula = f"=C{first}*B{first}/100"

    # The claim the message makes, as a formula over the products above it.
    count_row = last_row + 2
    label = sheet.Cells(count_row, 1)
    label.NumberFormat = "@"
    label.Value = f"{D.C09C_MESSAGE_LINE} at {D.C09C_SEGMENT:.0f} mUSD or more"
    label.Font.Bold = True
    count = sheet.Cells(count_row, 5)
    count.Formula = (f'=SUMPRODUCT((D{first_row}:D{last_row}="'
                     f'{D.C09C_MESSAGE_LINE}")*(E{first_row}:E{last_row}>='
                     f'{D.C09C_SEGMENT}))')
    count.Font.Bold = True
    stated = sheet.Cells(count_row + 1, 1)
    stated.NumberFormat = "@"
    stated.Value = "stated in the message"
    stated.Font.Color = rgb(S.PAGE["footnote"])
    told = sheet.Cells(count_row + 1, 5)
    told.Value = D.C09C_MESSAGE_COUNT
    told.Interior.Color = rgb("#FFF2CC")
    told.Font.Color = rgb(S.PAGE["footnote"])

    # The curves: margin down one column, one column of net sales per level.
    curve_head = count_row + 3
    axis_x, axis_y = t.axes
    tag = sheet.Cells(curve_head, 1)
    tag.NumberFormat = "@"
    tag.Value = "Iso gross profit"
    tag.Font.Bold = True
    sheet.Cells(curve_head, 2).Value = "Margin"
    for j, level in enumerate(D.C09C_ISO_PROFIT):
        sheet.Cells(curve_head, 3 + j).Value = level
    lo, hi = curve_head + 1, curve_head + C09C_CURVE_STEPS
    # Spread across the axis, but starting where the highest curve leaves the
    # top of the plot rather than at zero, where it runs to infinity.
    start = D.C09C_ISO_PROFIT[-1] * 100.0 / axis_y.maximum
    margins = [start + (axis_x.maximum - start) * i / (C09C_CURVE_STEPS - 1)
               for i in range(C09C_CURVE_STEPS)]
    rng = sheet.Range(sheet.Cells(lo, 2), sheet.Cells(hi, 2))
    rng.NumberFormat = "0.00"
    rng.Value = tuple((m,) for m in margins)
    for j, level in enumerate(D.C09C_ISO_PROFIT):
        rng = sheet.Range(sheet.Cells(lo, 3 + j), sheet.Cells(hi, 3 + j))
        rng.NumberFormat = "0.00"
        rng.Formula = (f"=IF({level}*100/$B{lo}>{axis_y.maximum},NA(),"
                       f"{level}*100/$B{lo})")
    return {"gross_profit": gp_col, "count": count_row, "curves": (lo, hi),
            "first": first_row, "last": last_row,
            "blocks": sorted(rows.values())}


def _add_c09_curves(chart, sheet, extras: dict) -> None:
    """Each iso-profit curve as a line series with no markers."""
    lo, hi = extras["curves"]
    for j, level in enumerate(D.C09C_ISO_PROFIT):
        series = chart.SeriesCollection().NewSeries()
        series.Name = f"{level:.0f} mUSD"
        series.XValues = sheet.Range(sheet.Cells(lo, 2), sheet.Cells(hi, 2))
        series.Values = sheet.Range(sheet.Cells(lo, 3 + j), sheet.Cells(hi, 3 + j))
        series.MarkerStyle = XL_NONE
        series.Format.Line.Visible = MSO_TRUE
        series.Format.Line.ForeColor.RGB = rgb("#A6A6A6")
        series.Format.Line.Weight = 1.25
        series.Smooth = False
        series.HasDataLabels = False


def verify_c09(sheet, t: D.Template, layout: L.XYLayout, extras: dict) -> list[str]:
    """The three derived quantities must be formulas, and must be right."""
    problems: list[str] = []
    first, last = extras["first"], extras["last"]

    blocks = [b for b in extras["blocks"]]
    for r in [blocks[0][0], blocks[1][1], blocks[-1][1]]:
        cell = sheet.Cells(r, extras["gross_profit"])
        if not str(cell.Formula).startswith("="):
            problems.append(f"{layout.template_id}: gross profit on row {r} is "
                            f"typed, not net sales times margin")

    count = sheet.Cells(extras["count"], 5)
    if not str(count.Formula).startswith("="):
        problems.append(f"{layout.template_id}: the segment count is typed; it "
                        f"is the one number on the page the message depends on")
    drawn = sum(1 for p in t.points
                if p.group == D.C09C_MESSAGE_LINE
                and D.c09c_gross_profit(p) >= D.C09C_SEGMENT)
    if count.Value is None or abs(float(count.Value) - drawn) > 0.5:
        problems.append(f"{layout.template_id}: the sheet counts {count.Value} "
                        f"products in the segment where the points give {drawn}")

    lo, hi = extras["curves"]
    for r in (lo, hi):
        for j in range(len(D.C09C_ISO_PROFIT)):
            if not str(sheet.Cells(r, 3 + j).Formula).startswith("="):
                problems.append(f"{layout.template_id}: the curve at row {r} "
                                f"column {3 + j} is typed, not derived")
    return problems


# The rungs an axis step is allowed to stand on, per decade. IBCS's own
# choices all sit on this ladder - 5 and 0.25 and 0.5 - which is what lets a
# fitted axis reproduce them exactly on the shipped figures.
AXIS_LADDER = (1.0, 2.0, 2.5, 5.0)


def _step_neighbour(step: float, up: bool) -> float:
    """The next rung of the 1-2-2.5-5 ladder above or below `step`."""
    decade = 10.0 ** math.floor(math.log10(step))
    mantissa = step / decade
    rungs = [m * d for d in (decade / 10, decade, decade * 10)
             for m in AXIS_LADDER]
    rungs.sort()
    here = min(range(len(rungs)), key=lambda i: abs(rungs[i] - step))
    return rungs[min(here + 1, len(rungs) - 1)] if up else rungs[max(here - 1, 0)]


def fit_xy_axis(axis: D.Axis, values) -> D.Axis:
    """An axis that covers the data, keeping the step the template chose.

    The XY pair are the only charts in the library that *show* a value axis -
    gridlines, ticks and numbers - which is why they cannot use the scale block
    every other family uses. Dividing the data by a span is invisible when the
    axis is stripped and unreadable when it is not: the reader would see 0.25
    where the figures say 29.16%.

    So these two are fitted at build instead. The step is left where the
    template put it, because how finely a reader should be able to read the
    axis is a design decision rather than a fact about the data - and keeping
    it is what makes the fit reproduce IBCS's own axes exactly on the shipped
    figures. It only moves when the data has left the decade it was chosen for,
    and then it moves a rung at a time.

    The limit this accepts: it fits when the workbook is built, not when a cell
    is edited. Excel will not bind an axis bound to a formula, and the way
    round that is the very thing a visible axis rules out.
    """
    numbers = [v for v in values if v is not None]
    if not numbers:
        return axis
    lo, hi = min(min(numbers), axis.minimum), max(numbers)
    step = axis.step
    guard = 0
    while (hi - lo) / step > 8 and guard < 20:
        step, guard = _step_neighbour(step, up=True), guard + 1
    while (hi - lo) / step < 1.5 and guard < 20:
        step, guard = _step_neighbour(step, up=False), guard + 1
    return dc.replace(axis,
                      minimum=math.floor(lo / step) * step,
                      maximum=math.ceil(hi / step) * step,
                      step=step)


def fit_xy(template: D.Template) -> D.Template:
    """Both axes of an XY template, fitted to the points it carries."""
    x_axis, y_axis = template.axes
    return dc.replace(template, axes=(
        fit_xy_axis(x_axis, [p.x for p in template.points]),
        fit_xy_axis(y_axis, [p.y for p in template.points])))


def build_xy_sheet(excel, sheet, template: D.Template) -> tuple[list[str], list]:
    """One XY template onto one sheet: a row per point, then one chart."""
    layout = L.xy_layout_for(f"{template.id}{template.variant}")
    sheet.Name = layout.template_id
    template = fit_xy(template)

    _write_xy_inputs(sheet, template, layout)
    rows = _write_xy_data(sheet, template, layout)
    extras = (_write_c09_extras(sheet, template, layout, rows)
              if template.id == "C09" else None)
    excel.Calculate()

    left = sum(sheet.Columns(i).Width for i in range(1, 6)) + layout.chart_gap
    shapes = []
    top = layout.chart_top
    for field, size, bold, height in (("entity", 11, False, 15.0),
                                      ("subject", 11, True, 15.0),
                                      ("period", 11, False, 15.0),
                                      ("message", 10, False, 26.0)):
        shapes.append(_linked_textbox(
            sheet, layout, f"title_{layout.template_id}_{field}",
            layout.title_row(field), left, top, layout.chart_width, height,
            size=size, bold=bold))
        top += height
    obj = _add_xy_chart(sheet, template, layout, rows, left, extras)
    obj.Top = top + layout.chart_gap
    excel.Calculate()

    problems = verify_xy(sheet, template, layout, rows, [obj])
    if extras is not None:
        problems += verify_c09(sheet, template, layout, extras)
    sized = any(p.size is not None for p in template.points)
    print(f"{chr(10)}{layout.template_id}: {len(template.points)} "
          f"{'bubbles on one size ruler' if sized else 'points'} in "
          f"{len(layout.scenarios(template))} series; axes "
          f"{template.axes[0].minimum:g}..{template.axes[0].maximum:g} and "
          f"{template.axes[1].minimum:g}..{template.axes[1].maximum:g}"
          + ("" if extras is None else
             f", {len(D.C09C_ISO_PROFIT)} derived iso-profit curves"))
    sheet.Range("A1").Select()
    # Charts only: the export step calls .Chart on everything it is handed, and
    # a linked text box has none.
    return problems, [obj]


# --------------------------------------------------------------------------- #
# Tree sheets
# --------------------------------------------------------------------------- #
#
# Six small charts and the arithmetic between them. Three things make this
# sheet different from every other one in the workbook:
#
# * **Three of the six rows are formulas.** Return, net sales and invested
#   capital are typed; return on sales, capital turnover and ROI are quotients
#   of them. That is not a convenience - it is the template's subject. A tree
#   whose ratios were typed would keep drawing its connectors after someone
#   edited a base measure, and the picture would state a calculation that was
#   no longer being done.
#
# * **The scale belongs to the unit, not to the chart.** Excel's default is one
#   scale per chart, which would draw net sales' 27.7 and return's 5.5 the same
#   height. Each chart's value axis is therefore set from
#   ``points_per_unit[group]`` and its plot area locked to
#   ``(maximum - minimum) x points_per_unit``, so a kEUR is the same number of
#   points in all three currency boxes. ``verify_tree`` reads it back off the
#   built charts rather than trusting the assignment.
#
# * **The box heights differ and that is the consequence, not a contradiction.**
#   A box is as tall as its own range needs at the group's scale. Forcing the
#   six boxes to a common height is precisely the mistake the shared scale
#   exists to prevent.

MSO_SHAPE_OVAL = 9
MSO_AUTOSIZE_SHAPE_TO_FIT_TEXT = 1
XL_TICK_LABEL_NEXT_TO_AXIS = 4
# The box outline. Chart furniture rather than notation, so it is a
# constant here and in the SVG layout rather than a palette entry.
TREE_BOX_STROKE = "#7F7F7F"
TREE_CONNECTOR = "#404040"
TREE_TYPED = ("return", "net_sales", "capital")
# result -> (numerator, denominator, scale). The worksheet builds the same
# quotients the data layer declares in C11A_TREE.links; _check_c11a proves the
# two agree, and verify_tree proves the cells really hold formulas.
TREE_DERIVED = {"ros": ("return", "net_sales", 100.0),
                "turnover": ("net_sales", "capital", 1.0),
                "roi": ("return", "capital", 100.0)}
TREE_STEP = {"percent": 5.0, "kEUR": 2.5, "turnover": 0.5}
TREE_PAD = 0.14          # headroom for the value labels, as a share of range
TREE_TOP_PAD = 24.0      # subtitle
TREE_BOTTOM_PAD = 30.0   # the two category rows
TREE_PLOT_LEAD = 8.0
TREE_PLOT_TRAIL = 8.0


def _tree_bounds(values, group: str) -> tuple[float, float]:
    """Axis minimum and maximum for one box: the data, padded and rounded.

    Zero is always included. A column chart measures from the axis, so an axis
    that started at 22 would draw net sales' 22.1 as a sliver and its 27.7 as
    five times the sliver - a fivefold difference where the data has none.
    """
    step = TREE_STEP[group]
    lo, hi = min(0.0, min(values)), max(0.0, max(values))
    pad = TREE_PAD * (hi - lo)
    return (math.floor((lo - pad) / step) * step,
            math.ceil((hi + pad) / step) * step)


def _write_tree_data(sheet, t: D.Template, layout: L.TreeLayout) -> dict:
    """Years, scenarios, the two label rows, three typed measures, three formulas."""
    sheet.Columns(1).ColumnWidth = layout.label_width
    for i in range(2, 2 + len(t.categories)):
        sheet.Columns(i).ColumnWidth = layout.value_width

    head = layout.first_row
    n = len(t.categories)
    for j, name in enumerate(t.categories):
        cell = sheet.Cells(head, 2 + j)
        cell.NumberFormat = "@"
        cell.Value = name
        cell.HorizontalAlignment = XL_RIGHT_ALIGN
        scen = sheet.Cells(head + 1, 2 + j)
        scen.NumberFormat = "@"
        scen.Value = t.category_scenarios[j]
        scen.Font.Color = rgb(S.PAGE["footnote"])
    sheet.Cells(head, 1).Value = "Year"
    sheet.Cells(head + 1, 1).Value = "Scenario"

    # The two rows the category axis actually reads. Both are formulas, so the
    # suppression rules are stated on the sheet rather than buried in a
    # renderer: a year prints only where the template says to print one, and a
    # scenario prints only where it changes.
    # Scenario first, year second. Excel draws the *last* row of a multi-level
    # category range nearest the axis, so listing the year first put "AC" up
    # against the zero line with "2021" underneath it - the reference has them
    # the other way round.
    sheet.Cells(head + 2, 1).Value = "Axis scenario"
    sheet.Cells(head + 3, 1).Value = "Axis year"
    printed = set(t.tree.printed_categories)
    for k in range(n):
        col = L.col_letter(2 + k)
        year = sheet.Cells(head + 3, 2 + k)
        # No "@" on a cell that is about to be given a formula. A text format
        # applied *first* makes Excel store the formula as literal text, and
        # the category axis then reads out "=B10" in place of the year - which
        # renders perfectly and is nonsense. The format goes on afterwards, or
        # not at all.
        year.Formula = f"={col}{head}" if k in printed else '=""'
        year.HorizontalAlignment = XL_RIGHT_ALIGN
        scen = sheet.Cells(head + 2, 2 + k)
        if k == 0:
            scen.Formula = f"={col}{head + 1}"
        else:
            prev = L.col_letter(1 + k)
            scen.Formula = (f'=IF({col}{head + 1}={prev}{head + 1},"",'
                            f'{col}{head + 1})')
        scen.HorizontalAlignment = XL_RIGHT_ALIGN

    rows: dict[str, int] = {}
    r = head + 4
    for key in TREE_TYPED:
        rows[key] = r
        sheet.Cells(r, 1).Value = t.tier(key).label
        series = t.tier(key)
        for j in range(n):
            cell = sheet.Cells(r, 2 + j)
            cell.NumberFormat = layout.number_format
            value = next((s.values[j] for s in series.series
                          if s.values[j] is not None), None)
            cell.Value = value
            cell.Interior.Color = rgb("#FFF2CC")
        r += 1
    for key, (num, den, scale) in TREE_DERIVED.items():
        rows[key] = r
        sheet.Cells(r, 1).Value = t.tier(key).label
        for j in range(n):
            cell = sheet.Cells(r, 2 + j)
            cell.NumberFormat = layout.number_format
            col = L.col_letter(2 + j)
            factor = "" if scale == 1.0 else f"*{scale:g}"
            cell.Formula = f"={col}{rows[num]}/{col}{rows[den]}{factor}"
        r += 1
    return rows


def _add_tree_chart(sheet, t: D.Template, layout: L.TreeLayout,
                    node: D.TreeNode, rows: dict, shift: int, spans: dict,
                    head: int, left: float, top: float):
    """One box of the tree, on its group's scale."""
    tier = t.tier(node.key)
    n = len(t.categories)
    values = _tree_values(t, node.key)
    low, high = _tree_bounds(values, node.scale_group)
    per_unit = layout.scale_of(node.scale_group)
    # The box's height comes from the *raw* range, so a box whose numbers span
    # more is taller - which is how boxes sharing a unit come out sharing a
    # scale. Only the axis is divided; the geometry is untouched.
    plot_height = (high - low) * per_unit
    span = sheet.Range(spans[node.scale_group]).Value
    if not span:
        raise ValueError(f"{layout.template_id}: the {node.scale_group} span is "
                         f"empty, so every box in it would divide by nothing")
    height = plot_height + TREE_TOP_PAD + TREE_BOTTOM_PAD

    obj = sheet.ChartObjects().Add(left, top, layout.box_width, height)
    obj.Name = f"tree_{node.key}"
    obj.Placement = XL_FREE_FLOATING
    chart = obj.Chart
    chart.ChartType = XL_COLUMN_CLUSTERED

    r = rows[node.key]
    # The mirror is on hidden rows, which Excel will not plot from unless told.
    chart.PlotVisibleOnly = False
    series = chart.SeriesCollection().NewSeries()
    series.Name = tier.label
    series.Values = sheet.Range(sheet.Cells(r + shift, 2),
                                sheet.Cells(r + shift, 1 + n))
    # Two rows, so Excel draws a two-level category axis: the year, and the
    # scenario under the year it starts at. The integrated legend, from cells.
    # One level, the year. A two-row range gives Excel a multi-level category
    # axis, which looks right and places the scenario wrong: Excel *centres* an
    # outer group over the periods it spans, so "AC" came out under 2023 and
    # "PL" between 2026 and 2027, where the reference puts each under the first
    # period it applies to. It also draws a boxed separator grid around the
    # labels. Both go; _tree_scenario_labels puts the scenario back as two
    # cell-linked boxes.
    series.XValues = sheet.Range(sheet.Cells(head + 3, 2),
                                 sheet.Cells(head + 3, 1 + n))
    chart.ChartGroups(1).GapWidth = (layout.gap_width if node.column < 2
                                     else layout.wide_gap_width)

    strip_chrome(chart, show_categories=True)
    cat_axis = chart.Axes(XL_CATEGORY)
    # Next to the axis, not below the plot: IBCS puts the year at the zero line
    # so a negative column runs past it rather than pushing it down the page.
    cat_axis.TickLabelPosition = XL_TICK_LABEL_NEXT_TO_AXIS
    cat_axis.Format.Line.Visible = MSO_FALSE
    chart.ChartArea.Format.Line.Visible = MSO_TRUE
    chart.ChartArea.Format.Line.ForeColor.RGB = rgb(TREE_BOX_STROKE)
    chart.ChartArea.Format.Line.Weight = 0.75

    axis = chart.Axes(XL_VALUE)
    axis.MinimumScale = low / span
    axis.MaximumScale = high / span

    series.HasDataLabels = True
    labels = series.DataLabels()
    labels.NumberFormat = layout.number_format
    for j in range(n):
        try:
            series.Points(j + 1).DataLabel.Formula = (
                f"='{sheet.Name}'!{a1_abs(r + 2 * shift, 2 + j)}")
        except Exception:                                     # noqa: BLE001
            pass
    labels.Font.Size = 9
    labels.Font.Name = "Arial"
    # OutsideEnd, not Above. On a column chart Above and Below raise a bare
    # "Exception occurred" naming nothing - they are line and XY positions.
    # OutsideEnd means "past the end of the bar", which already follows the
    # sign: it sits over a positive column and under a negative one, which is
    # exactly what the reference draws and saves positioning any point by hand.
    labels.Position = XL_LABEL_OUTSIDE_END

    for j in range(n):
        point = series.Points(j + 1)
        scenario = t.category_scenarios[j]
        apply_scenario(point.Format, scenario)
        if values[j] < 0:
            # The lighter tone of the scenario's pair, as the reference draws
            # it - see the SVG renderer for why a negative takes the waterfall's
            # subtracting shade.
            point.Format.Fill.ForeColor.RGB = rgb(S.waterfall_fill(scenario, -1))
            point.Format.Line.Visible = MSO_FALSE

    # Last, because adding the series and then relaying the chart out brings
    # Excel's automatic single-series title back after strip_chrome removed it.
    chart.HasTitle = False
    _fit_tree_plot(obj, layout, plot_height)
    return obj, (low, high, plot_height, r)


def _fit_tree_plot(obj, layout: L.TreeLayout, plot_height: float,
                   attempts: int = 8) -> None:
    """Grow the box until its plot area really is ``plot_height`` points tall.

    A visible category axis *caps* InsideHeight, and does it silently: ask a
    166pt chart for a 112pt plot with two rows of tick labels under it and the
    assignment neither raises nor takes - it comes back 71pt, with InsideTop
    driven to -4. Excel is reserving whatever the labels need and shrinking the
    plot to fit, and no amount of re-locking changes that, because nothing
    failed.

    So the chart is grown by the shortfall and re-measured, which is the same
    move the table sheets make to align a chart to a row band. The label
    reservation is not knowable in advance - it depends on the font, the number
    of axis levels and the longest label - but it is stable, so two or three
    passes converge.

    This is what keeps the scale exact. Accepting the clamped height would put
    every box on its own px-per-unit and quietly undo the one rule the template
    exists to demonstrate.
    """
    for _ in range(attempts):
        _lock_tree_plot(obj.Chart, layout, plot_height)
        actual = obj.Chart.PlotArea.InsideHeight
        shortfall = plot_height - actual
        if abs(shortfall) < 0.25:
            return
        obj.Height = obj.Height + shortfall
    _lock_tree_plot(obj.Chart, layout, plot_height)


def _lock_tree_plot(chart, layout: L.TreeLayout, plot_height: float) -> None:
    """Pin the plot area so the group's scale survives Excel's autolayout.

    Same retry dance as ``lock_plot_area`` and for the same reasons - assign
    height before top, tolerate all but the last failure, re-fetch PlotArea
    each pass - but with the numbers passed in rather than read off a
    SheetLayout, because a tree has no shared category axis to align to.
    """
    attempts = 4
    for attempt in range(attempts):
        try:
            plot = chart.PlotArea
            plot.InsideWidth = layout.box_width - TREE_PLOT_LEAD - TREE_PLOT_TRAIL
            plot.InsideLeft = TREE_PLOT_LEAD
            plot.InsideHeight = plot_height
            plot.InsideTop = TREE_TOP_PAD
        except Exception:                                     # noqa: BLE001
            if attempt == attempts - 1:
                raise


def _tree_subtitle(sheet, t: D.Template, node: D.TreeNode, obj) -> list:
    """The measure name bold, its unit plain, over the top of the box.

    Two boxes rather than one with a bold run, because ``TextRange2.Characters``
    is not reachable through pywin32 - it answers "Does not support a
    collection" - so a mixed-weight line cannot be built inside a single frame.
    The bold box is auto-sized to its text and the plain box starts where it
    ends, which measures the gap rather than estimating it.
    """
    tier = t.tier(node.key)
    boxes = []
    left = obj.Left + 9
    for text, bold in ((tier.label, True),
                       (f" in {node.unit}" if node.unit else "", False)):
        if not text:
            continue
        box = sheet.Shapes.AddTextbox(MSO_TEXT_HORIZONTAL, left, obj.Top + 4,
                                      60, 14)
        box.Name = f"subtitle_{node.key}_{'m' if bold else 'u'}"
        box.Line.Visible = MSO_FALSE
        box.Fill.Visible = MSO_FALSE
        frame = box.TextFrame2
        frame.MarginLeft = frame.MarginRight = 0
        frame.MarginTop = frame.MarginBottom = 0
        frame.WordWrap = MSO_FALSE
        frame.TextRange.Text = text
        frame.TextRange.Font.Size = 10
        frame.TextRange.Font.Name = "Arial"
        frame.TextRange.Font.Bold = MSO_TRUE if bold else MSO_FALSE
        # Auto-size, then read the width back: the unit has to start exactly
        # where the measure name stops, and Excel is the only thing that knows
        # how wide it drew the text.
        frame.AutoSize = MSO_AUTOSIZE_SHAPE_TO_FIT_TEXT
        box.Placement = XL_FREE_FLOATING
        left = box.Left + box.Width
        boxes.append(box)
    return boxes


def _tree_zero_and_split(sheet, t: D.Template, node: D.TreeNode, obj,
                         bounds) -> list:
    """The zero line and the plan rule, from the plot area as Excel laid it out.

    Both are shapes rather than chart furniture, and both are measured off
    ``PlotArea`` rather than computed from the numbers that were asked for.
    The category axis line has to be switched off - it is the same object that
    draws the separator grid a multi-level axis puts between its labels, and
    IBCS wants the labels bare - so the zero line has to be drawn back.
    """
    low, high, _plot_height, _row = bounds
    chart = obj.Chart
    plot = chart.PlotArea
    inside_top, inside_h = plot.InsideTop, plot.InsideHeight
    inside_left, inside_w = plot.InsideLeft, plot.InsideWidth
    zero_y = obj.Top + inside_top + inside_h * high / (high - low)

    shapes = []
    rule = sheet.Shapes.AddLine(obj.Left + inside_left, zero_y,
                                obj.Left + inside_left + inside_w, zero_y)
    rule.Name = f"zero_{node.key}"
    rule.Line.ForeColor.RGB = rgb("#000000")
    rule.Line.Weight = 0.75
    rule.Placement = XL_FREE_FLOATING
    shapes.append(rule)

    n = len(t.categories)
    after = next(j for j in range(1, n)
                 if t.category_scenarios[j] != t.category_scenarios[j - 1])
    x = obj.Left + inside_left + inside_w * after / n
    split = sheet.Shapes.AddLine(x, zero_y - 48.0, x, zero_y + 20.0)
    split.Name = f"split_{node.key}"
    split.Line.ForeColor.RGB = rgb("#000000")
    split.Line.Weight = 0.9
    split.Placement = XL_FREE_FLOATING
    shapes.append(split)
    return shapes


def _tree_connectors(sheet, t: D.Template, layout: L.TreeLayout,
                     placed: dict) -> list:
    """The spines, the stubs and the operator circles.

    Drawn from ``t.tree.links`` rather than from coordinates, exactly as the SVG
    does, so the two engines cannot disagree about what the tree claims. Each
    stub meets a box at that box's vertical centre and each circle sits at the
    centre of the box its link produces.
    """
    shapes = []
    radius = 9.0

    def centre(key: str) -> float:
        obj = placed[key]
        return obj.Top + obj.Height / 2

    for link in t.tree.links:
        result = placed[link.result]
        spine_x = result.Left + result.Width + layout.column_gap / 2
        ys = []
        for k, operand in enumerate((link.left, link.right)):
            box = placed[operand]
            # Net sales feeds two links; its two stubs are pulled apart so the
            # second does not lie on top of the first.
            offset = 0.0
            if sum(1 for l in t.tree.links
                   if operand in (l.left, l.right)) > 1:
                offset = -6.0 if link.result == "ros" else 6.0
            y = centre(operand) + offset
            line = sheet.Shapes.AddLine(spine_x, y, box.Left, y)
            line.Line.ForeColor.RGB = rgb(TREE_CONNECTOR)
            line.Line.Weight = 1.0
            line.Placement = XL_FREE_FLOATING
            shapes.append(line)
            ys.append(y)
        spine = sheet.Shapes.AddLine(spine_x, min(ys), spine_x, max(ys))
        spine.Line.ForeColor.RGB = rgb(TREE_CONNECTOR)
        spine.Line.Weight = 1.0
        spine.Placement = XL_FREE_FLOATING
        shapes.append(spine)

        cy = centre(link.result)
        stub = sheet.Shapes.AddLine(result.Left + result.Width, cy,
                                    spine_x - radius, cy)
        stub.Line.ForeColor.RGB = rgb(TREE_CONNECTOR)
        stub.Line.Weight = 1.0
        stub.Placement = XL_FREE_FLOATING
        shapes.append(stub)

        circle = sheet.Shapes.AddShape(MSO_SHAPE_OVAL, spine_x - radius,
                                       cy - radius, radius * 2, radius * 2)
        circle.Name = f"op_{link.result}"
        circle.Fill.ForeColor.RGB = rgb("#FFFFFF")
        circle.Fill.Solid()
        circle.Line.ForeColor.RGB = rgb(TREE_CONNECTOR)
        circle.Line.Weight = 1.0
        frame = circle.TextFrame2
        frame.TextRange.Text = link.glyph
        frame.TextRange.Font.Size = 10
        frame.TextRange.Font.Name = "Arial"
        frame.TextRange.Font.Fill.ForeColor.RGB = rgb(TREE_CONNECTOR)
        frame.TextRange.Font.Bold = MSO_FALSE
        circle.Placement = XL_FREE_FLOATING
        shapes.append(circle)
    return shapes


def _tree_values(t: D.Template, key: str) -> list[float]:
    """One box's series, with the AC and PL halves merged back together."""
    tier = t.tier(key)
    return [next(v for v in (s.values[j] for s in tier.series) if v is not None)
            for j in range(len(t.categories))]


def verify_tree(t: D.Template, placed: dict, built: dict,
                layout: L.TreeLayout, spans: dict | None = None) -> list[str]:
    """Prove the boxes really are on their group's scale, and the ratios derived.

    Stated as a height rather than as a ratio, because a height is what a reader
    can be misled by: take the largest value in a group and ask how many points
    it draws in each box of that group. If two boxes disagree by more than a
    point, someone carrying a figure from one to the other reads it wrong.
    Comparing rounded points-per-unit instead fails on a third decimal no eye
    could resolve, which tests the arithmetic rather than the picture.

    Both halves of the template are checked. A chart left on its own maximum
    draws whatever its tallest column happens to be at full height, so return's
    5.5 and net sales' 27.7 come out the same size; and a ratio typed rather
    than computed keeps its old value when a base measure moves, which makes
    the connectors state a calculation nobody is doing.
    """
    problems: list[str] = []
    rates: dict[str, list[tuple[str, float]]] = {}
    sheet0 = placed[t.tree.nodes[0].key].Parent
    for node in t.tree.nodes:
        chart = placed[node.key].Chart
        axis = chart.Axes(XL_VALUE)
        # The axis runs in fractions of the group's span, so the height it
        # covers is points per *fraction*. Multiplying back by the span is what
        # turns it into points per unit - which is the number a reader is being
        # promised is the same across the group.
        divisor = axis.MaximumScale - axis.MinimumScale
        if spans:
            divisor *= sheet0.Range(spans[node.scale_group]).Value
        rates.setdefault(node.scale_group, []).append(
            (node.key, chart.PlotArea.InsideHeight / divisor))

    for group, entries in rates.items():
        want = layout.scale_of(group)
        biggest = max(abs(v) for key, _ in entries for v in _tree_values(t, key))
        drawn = {key: biggest * rate for key, rate in entries}
        spread = max(drawn.values()) - min(drawn.values())
        if spread > 1.0:
            detail = ", ".join(f"{k} {v:.1f}pt" for k, v in drawn.items())
            problems.append(
                f"the {group} boxes draw {biggest:g} at different heights "
                f"({detail}) - a figure carried between them would be read wrong")
        for key, rate in entries:
            if abs(rate - want) / want > 0.02:
                problems.append(
                    f"{key} draws {rate:.3f}pt per unit, not the {want:.3f} its "
                    f"{group} group shares")

    # The three ratio rows must be formulas. A tree drawing its connectors over
    # typed quotients states a calculation it is not doing.
    sheet = placed[t.tree.nodes[0].key].Parent
    for key in TREE_DERIVED:
        row = built[key][3]
        if not str(sheet.Cells(row, 2).Formula).startswith("="):
            problems.append(
                f"{key} is typed, not derived - edit a base measure and this box "
                f"would keep its old value while the connectors still claim it "
                f"is computed")
    return problems


def _tree_scenario_labels(sheet, t: D.Template, layout: L.TreeLayout,
                          node: D.TreeNode, obj, head: int, bounds) -> list:
    """AC and PL, each under the first period it applies to.

    Cell-linked, so the integrated legend is still something the sheet says
    rather than something the builder decided: the cells they point at hold
    ``=IF(this period's scenario = the previous one, "", it)``, which is the
    suppression rule itself, written down.
    """
    plot = obj.Chart.PlotArea
    inside_left, inside_w = plot.InsideLeft, plot.InsideWidth
    inside_top, inside_h = plot.InsideTop, plot.InsideHeight
    n = len(t.categories)
    slot = inside_w / n
    boxes = []
    for j in range(n):
        if j and t.category_scenarios[j] == t.category_scenarios[j - 1]:
            continue
        box = sheet.Shapes.AddTextbox(
            MSO_TEXT_HORIZONTAL,
            obj.Left + inside_left + slot * j, 0, slot, 13)
        box.Name = f"scenario_{node.key}_{j}"
        box.Line.Visible = MSO_FALSE
        box.Fill.Visible = MSO_FALSE
        frame = box.TextFrame2
        frame.MarginLeft = frame.MarginRight = 0
        frame.MarginTop = frame.MarginBottom = 0
        frame.WordWrap = MSO_FALSE
        try:
            box.DrawingObject.Formula = (
                f"='{sheet.Name}'!${L.col_letter(2 + j)}${head + 2}")
        except Exception:                                     # noqa: BLE001
            frame.TextRange.Text = t.category_scenarios[j]
        frame.TextRange.Font.Size = 9
        frame.TextRange.Font.Name = "Arial"
        frame.TextRange.ParagraphFormat.Alignment = 2          # centred in slot
        # Measured down from the *zero line*, not from the bottom of the plot.
        # The year labels sit next to the axis, and in a box with negative
        # columns the axis is somewhere in the middle - so anchoring to the
        # plot bottom edge left "AC" stranded a hundred points below the year
        # it belongs under. Where the axis minimum is zero the two coincide,
        # which is why the small boxes looked right and the tall ones did not.
        low, high, _plot_height, _row = bounds
        zero_y = obj.Top + inside_top + inside_h * high / (high - low)
        box.Top = zero_y + 25.0
        box.Placement = XL_FREE_FLOATING
        boxes.append(box)
    return boxes


def _tree_annotations(sheet, t: D.Template, placed: dict, built: dict) -> list:
    """The highlight oval, round the one value the message is about.

    Positioned from the plot area rather than from the chart, so it lands on
    the value label wherever Excel finally put the plot - the same rule the
    zero line and the scenario labels follow.
    """
    shapes = []
    for a in t.annotations:
        if a.kind != "oval":
            continue
        key, index = a.target
        obj = placed[key]
        low, high, _plot_height, _row = built[key]
        plot = obj.Chart.PlotArea
        inside_h, inside_top = plot.InsideHeight, plot.InsideTop
        inside_left, inside_w = plot.InsideLeft, plot.InsideWidth
        n = len(t.categories)
        value = _tree_values(t, key)[index]
        zero_y = obj.Top + inside_top + inside_h * high / (high - low)
        y = zero_y - value * inside_h / (high - low)
        cx = obj.Left + inside_left + inside_w * (index + 0.5) / n
        rx, ry = 17.0, 9.0
        oval = sheet.Shapes.AddShape(MSO_SHAPE_OVAL, cx - rx, y - 9.0 - ry,
                                     rx * 2, ry * 2)
        oval.Name = f"highlight_{key}_{index}"
        oval.Fill.Visible = MSO_FALSE
        oval.Line.ForeColor.RGB = rgb(S.ANNOTATION["highlight"])
        oval.Line.Weight = 1.25
        oval.Placement = XL_FREE_FLOATING
        shapes.append(oval)
    return shapes


def build_tree_sheet(excel, sheet, template: D.Template) -> tuple[list[str], list]:
    """One driver tree onto one sheet: the data block, six charts, the links.

    Built in passes, because two of the steps change geometry the next one
    depends on. ``_fit_tree_plot`` grows a box until its plot area is the height
    the shared scale needs - net sales ends 60pt taller than it was asked for -
    so nothing can be stacked until every box has been through it. And the zero
    line and plan rule are placed from ``PlotArea`` as Excel finally laid it
    out, so they come last of all.
    """
    layout = L.tree_layout_for(f"{template.id}{template.variant}")
    sheet.Name = layout.template_id

    _write_input_block(sheet, template, layout, "PL")
    rows = _write_tree_data(sheet, template, layout)
    # One span per unit. The boxes of a driver tree are measured in three
    # different things - per cent, a turnover ratio and currency - and the
    # standard's claim is that boxes sharing a unit share a ruler. One span
    # cell per unit is what enforces that: they cannot drift apart because
    # there is only one number for each of them to divide by.
    tree_shift, tree_spans = _write_row_scale(
        sheet, layout, rows,
        [(node.key, layout.number_format, node.scale_group)
         for node in template.tree.nodes],
        len(template.categories))
    excel.Calculate()

    head = layout.first_row
    park = data_zone_width(sheet, layout) + layout.chart_left
    title_top = layout.chart_top
    layout = dc.replace(layout, chart_top=layout.chart_top + TITLE_BLOCK_HEIGHT)

    # Pass 1: build and fit. Position is provisional - a fitted box is taller
    # than the one that was asked for, so where it ends up cannot be known yet.
    placed, built = {}, {}
    for node in template.tree.nodes:
        obj, bounds = _add_tree_chart(sheet, template, layout, node, rows,
                                      tree_shift, tree_spans,
                                      head, park, layout.chart_top)
        placed[node.key] = obj
        built[node.key] = bounds

    # Pass 2: stack each tree column from the heights the boxes actually came
    # out at, then centre the stacks against the tallest so the tree reads as a
    # tree rather than as three lists.
    columns: dict[int, list[D.TreeNode]] = {}
    for node in template.tree.nodes:
        columns.setdefault(node.column, []).append(node)
    totals = {c: sum(placed[n.key].Height for n in ns)
                 + layout.row_gap * (len(ns) - 1)
              for c, ns in columns.items()}
    tallest = max(totals.values())

    left = park
    for col in sorted(columns):
        top = layout.chart_top + (tallest - totals[col]) / 2
        for node in columns[col]:
            obj = placed[node.key]
            obj.Left, obj.Top = left, top
            top += obj.Height + layout.row_gap
        left += layout.box_width + layout.column_gap

    # Pass 3: re-assert what Excel's relayouts undo. Moving a chart and
    # resizing it both bring the automatic single-series title back, and both
    # can disturb a plot area that was locked before the move.
    for node in template.tree.nodes:
        chart = placed[node.key].Chart
        chart.HasTitle = False
        _fit_tree_plot(placed[node.key], layout, built[node.key][2])
        chart.HasTitle = False

    # Pass 4: everything positioned from the finished plot geometry.
    for node in template.tree.nodes:
        _tree_subtitle(sheet, template, node, placed[node.key])
        _tree_zero_and_split(sheet, template, node, placed[node.key],
                             built[node.key])
        _tree_scenario_labels(sheet, template, layout, node,
                              placed[node.key], head, built[node.key])
    _tree_connectors(sheet, template, layout, placed)
    _tree_annotations(sheet, template, placed, built)
    add_title_block(sheet, template, layout, park, title_top,
                    (layout.box_width + layout.column_gap) * 3)

    problems = verify_tree(template, placed, built, layout,
                           tree_spans)
    print(f"\n{layout.template_id}: {len(placed)} boxes in "
          f"{len(columns)} tree columns")
    for group, keys in template.tree.groups().items():
        print(f"    {group:<9} {layout.scale_of(group):7.3f}pt per unit  "
              f"{', '.join(keys)}")
    sheet.Range("A1").Select()
    return problems, list(placed.values())


# --------------------------------------------------------------------------- #
# Panel sheets
# --------------------------------------------------------------------------- #
#
# C13 is a panel chart, and Excel has no panel chart type. The workaround
# everyone reaches for - build one chart and copy it sixteen times - leaves the
# reader with sixteen chart objects whose axes drift apart the first time
# anyone edits a figure, which is the one thing a small multiple may not do.
#
# So the grid is not built here. The `panel-charts` skill builds the whole
# 4x4 as a **single native chart**, and this sheet is its first caller: it
# supplies the notation - IBCS greys, the variance colours, a hollow head for a
# planned period, the AC|PL rule - and the panel engine supplies the geometry.
# Everything below is either the IBCS half of that bargain or the wiring
# between the two.
#
# The workbook stays live across the seam. The engine's input block is what its
# formulas fan out into the plot table, so after the grid is built every cell
# of that block is **overwritten with a formula** reading the absolutes typed
# above it. Retype one location's net profit and the location average moves,
# every one of the fifteen panels' variances moves with it, and every pin on
# the page follows - which is what `check_panels_are_live` proves.

PANEL_INPUT_FILL = "#FFF2CC"


def _c13_panel_style(t: D.Template):
    """The notation, handed to the panel engine as data rather than as code."""
    from panel_excel import PanelStyle                      # noqa: PLC0415

    good = _hex_to_triple(S.VARIANCE["good"])
    bad = _hex_to_triple(S.VARIANCE["bad"])
    ac = _hex_to_triple(S.scenario_fill("AC").fill)
    return PanelStyle(
        element_colours=[ac],
        # Green up and red down, from the palette rather than from the engine's
        # own blue-and-orange. A pin's colour is not decoration here: it says
        # which way the variance went, and it is the only channel that does.
        sign_colours=[good, bad],
        # Ten measured years then two planned, in every panel. The engine has
        # no idea what a scenario is - it takes a list of names and a table of
        # marks, which is exactly the seam that keeps it notation-neutral.
        point_scenarios=list(t.category_scenarios),
        scenario_marks={"AC": (ac, ac), "PL": (None, ac)},
        value_labels=True,
        label_format="+0;-0;0",
        # No chart title and no tick labels: the sheet carries the IBCS title
        # block above, and a panel that prints every point's own value has no
        # use for a value axis - which is the whole reason the axis is hidden.
        chart_title=False,
        value_ticks=False,
        label_size=7,
        marker_size=4,
        stem_weight=1.25,
        divider_after=t.category_scenarios.index("PL"),
        divider_colour=(0, 0, 0),
        divider_weight=0.9,
    )


def _hex_to_triple(value: str) -> tuple:
    value = value.lstrip("#")
    return tuple(int(value[i:i + 2], 16) for i in (0, 2, 4))


def _write_panel_source(sheet, t: D.Template, layout: L.PanelLayout,
                        grid: D.PanelGrid, first_row: int) -> dict:
    """The absolutes the reader edits, and the average derived from them.

    This is the sheet's real input. What the panel engine reads is a block of
    variances, and a variance is not something anyone types - it is what the
    page is *about*, so it has to be computed here or the chart stops being a
    claim about the data.
    """
    head = first_row
    sheet.Cells(head, 1).Value = "Location"
    for j, year in enumerate(t.categories):
        cell = sheet.Cells(head, 2 + j)
        cell.NumberFormat = "@"
        cell.Value = year
        cell.HorizontalAlignment = XL_RIGHT_ALIGN
        scen = sheet.Cells(head + 1, 2 + j)
        scen.NumberFormat = "@"
        scen.Value = t.category_scenarios[j]
        scen.Font.Color = rgb(S.PAGE["footnote"])
    sheet.Cells(head + 1, 1).Value = "Scenario"

    rows = {}
    r = head + 2
    for key, label, values in D.C13D_PANELS:
        rows[key] = r
        sheet.Cells(r, 1).Value = label
        for j, value in enumerate(values):
            cell = sheet.Cells(r, 2 + j)
            cell.NumberFormat = layout.number_format
            cell.Value = value
            cell.Interior.Color = rgb(PANEL_INPUT_FILL)
        r += 1

    # The reference series, and it is a formula. IBCS titles this panel
    # "average 25 locations"; it is the mean of the fifteen panels above, one
    # of which stands for eleven locations on its own. Dividing the same total
    # by 25 reproduces none of the twelve figures the reference prints, and by
    # 15 reproduces all twelve - so the formula says 15, and _check_c13d is
    # where that is argued rather than asserted.
    rows["average"] = r
    sheet.Cells(r, 1).Value = t.tier("average").label
    first, last = head + 2, r - 1
    for j in range(len(t.categories)):
        col = L.col_letter(2 + j)
        cell = sheet.Cells(r, 2 + j)
        cell.NumberFormat = layout.number_format
        cell.Formula = f"=AVERAGE({col}{first}:{col}{last})"
    sheet.Cells(r, 1).Font.Bold = True
    return rows


def _link_panel_input(sheet, t: D.Template, grid: D.PanelGrid, builder,
                      rows: dict) -> None:
    """Replace the engine's typed input block with the variance formulas.

    The panel engine is handed values so that it can size its shared scale, and
    then every cell it wrote is replaced by the arithmetic that produced them.
    Nothing about the grid changes - the formulas fan out through exactly the
    same plot table - but the sheet stops being a picture of a calculation and
    starts being the calculation.
    """
    average = rows["average"]
    order = [cell.key for cell in sorted(grid.cells,
                                         key=lambda c: (c.row, c.col))]
    for p, key in enumerate(order):
        r = builder.in_r0 + p
        for j in range(len(t.categories)):
            col = L.col_letter(2 + j)
            cell = sheet.Cells(r, builder.in_c0 + j)
            if key is None:
                # The sixteenth cell of a fifteen-panel grid. Left genuinely
                # empty, not zeroed: a zero is a variance of none, which is a
                # statement, and there is no location here to make it about.
                cell.ClearContents()
                continue
            cell.Formula = (f"=({col}{rows[key]}-{col}{average})"
                            f"/{col}{average}*100")
            cell.NumberFormat = "+0;-0;0"


def _reference_panel(sheet, t: D.Template, layout: L.PanelLayout,
                     grid: D.PanelGrid, builder, chart, rows: dict):
    """The location average, drawn into the one cell the pin grid leaves empty.

    A separate chart object, and that is the point rather than a compromise.
    The fifteen pin panels are variances in *per cent* and share one ruler -
    which is the whole argument of a small multiple. This panel is the average
    itself, in mEUR. There is no scale the two can share, so putting it inside
    the same chart would mean two units on one axis, which is exactly what the
    grid exists to prevent. Different measure, different ruler, own object.

    It is still notation: actual solid, planned outlined, on the shaded ground
    IBCS uses to mark the panel a page is measured against.
    """
    spec = builder.spec
    plot = chart.Chart.PlotArea
    left, top = chart.Left + plot.InsideLeft, chart.Top + plot.InsideTop
    width, height = plot.InsideWidth, plot.InsideHeight

    cell = next(c for c in grid.cells if c.key is None)
    block, band = spec.place(cell.row, cell.col)

    # The cell's rectangle, from the same slot arithmetic the grid is drawn by,
    # so it lands on the grid rather than near it.
    x0 = left + width * (spec.block_start(block) - 1.0) / spec.n_slots
    x1 = left + width * spec.block_end(block) / spec.n_slots
    y1 = top + height * (1.0 - band / spec.bands)
    y0 = top + height * (1.0 - (band + spec.band_frac) / spec.bands)

    obj = sheet.ChartObjects().Add(x0, y0, x1 - x0, y1 - y0)
    obj.Name = "reference_panel"
    obj.Placement = XL_FREE_FLOATING
    ch = obj.Chart
    ch.ChartType = XL_COLUMN_CLUSTERED
    ch.HasTitle = False
    ch.HasLegend = False
    ch.PlotVisibleOnly = False

    row = rows["average"]
    n = len(t.categories)
    series = ch.SeriesCollection().NewSeries()
    series.Values = sheet.Range(sheet.Cells(row, 2), sheet.Cells(row, 1 + n))
    series.XValues = sheet.Range(sheet.Cells(layout.first_row + 3, 2),
                                 sheet.Cells(layout.first_row + 3, 1 + n))
    ch.ChartGroups(1).GapWidth = 40

    for j, scenario in enumerate(t.category_scenarios, start=1):
        apply_scenario(series.Points(j).Format, scenario)

    series.HasDataLabels = True
    labels = series.DataLabels()
    labels.NumberFormat = "0"
    labels.Font.Size = 7
    labels.Font.Name = "Arial"
    labels.Position = XL_LABEL_OUTSIDE_END

    strip_chrome(ch, show_categories=False)
    ch.Axes(XL_VALUE).MinimumScale = 0
    # Headroom for the labels, which sit outside the column ends.
    ch.Axes(XL_VALUE).MaximumScale = max(D.C13D_AVERAGE) * 1.45
    ch.ChartArea.Format.Fill.Visible = MSO_TRUE
    ch.ChartArea.Format.Fill.Solid()
    ch.ChartArea.Format.Fill.ForeColor.RGB = rgb("#ECECEC")
    ch.ChartArea.Format.Line.Visible = MSO_FALSE

    title = sheet.Shapes.AddTextbox(MSO_TEXT_HORIZONTAL, x0 + 4, y0 + 1,
                                    x1 - x0 - 8, 12)
    title.Name = "reference_panel_title"
    title.Line.Visible = MSO_FALSE
    title.Fill.Visible = MSO_FALSE
    frame = title.TextFrame2
    frame.MarginLeft = frame.MarginRight = 0
    frame.MarginTop = frame.MarginBottom = 0
    frame.WordWrap = MSO_FALSE
    frame.TextRange.Text = t.tier("average").label
    frame.TextRange.Font.Size = 8
    frame.TextRange.Font.Name = "Arial"
    title.Placement = XL_FREE_FLOATING
    return obj


def build_panel_sheet(excel, sheet, template: D.Template) -> tuple[list[str], list]:
    """One small-multiples template onto one sheet, via the panel-charts skill."""
    try:
        import panel_spec as PS                             # noqa: PLC0415
        import panel_excel as PE                            # noqa: PLC0415
    except ImportError as exc:                              # noqa: BLE001
        raise SystemExit(chr(10).join((
            "error: C13D needs the companion `panel-charts` skill, which "
            "was not found (%s)." % exc,
            "  Install it beside this one, or set PANEL_CHARTS to its "
            "scripts directory.",
            "  The other sixteen templates do not need it - drop C13D "
            "from --template to build them."))) from exc

    layout = L.panel_layout_for(f"{template.id}{template.variant}")
    sheet.Name = layout.template_id
    grid = template.panel_grids[layout.grid]

    _write_input_block(sheet, template, layout, "AVG")
    rows = _write_panel_source(sheet, template, layout, grid, layout.first_row)
    excel.Calculate()

    order = sorted(grid.cells, key=lambda c: (c.row, c.col))
    names = [(c.label or (template.tier(c.key).label if c.key else ""))
             for c in order]
    values = []
    for cell in order:
        if cell.key is None:
            values.append([None] * len(template.categories))
        else:
            values.append([round(v, 4) for v in D.C13D_VARIANCE[cell.key]])

    spec = PS.PanelSpec(
        rows=grid.rows, cols=grid.cols, periods=len(template.categories),
        elements=1, kind="pin", sheet=sheet.Name,
        panel_names=names,
        period_labels=["'" + y[2:] for y in template.categories],
        element_names=["ΔØ %"],
        include_zero=True, nticks=3, title="")

    # The engine is given the whole sheet below the source block. It writes its
    # own input block, parameters, plot table and annotations there, and none
    # of it collides with the IBCS title block above because `origin` moves
    # every row it uses - which is the seam that makes this skill reusable
    # rather than a workbook generator.
    origin = rows["average"] + 3
    builder = PE.PanelBuilder(spec, values, _c13_panel_style(template))
    chart = builder.build(sheet, origin=origin, page_setup=False, intro=False)
    # Named, because the sheet now holds two chart objects and "Chart 1" tells
    # neither a checker nor a reader which one is the grid.
    chart.Name = "panel_grid"

    # The panel engine parks its chart below every row it wrote, so the title
    # goes in the gap above it - and the chart moves down to make that gap the
    # right size.
    chart.Top = chart.Top + TITLE_BLOCK_HEIGHT
    add_title_block(sheet, template, layout, chart.Left,
                    chart.Top - TITLE_BLOCK_HEIGHT, chart.Width)

    reference = _reference_panel(sheet, template, layout, grid, builder,
                                 chart, rows)

    _link_panel_input(sheet, template, grid, builder, rows)
    excel.Calculate()

    problems = builder.verify(chart.Chart)
    problems += verify_panel_grid(template, grid, builder, sheet, chart,
                                  extra_charts=1)

    print(f"\n{layout.template_id}: {len(grid.drawn())} panels in one "
          f"{grid.rows}x{grid.cols} chart object, "
          f"{chart.Chart.SeriesCollection().Count} series")
    sheet.Range("A1").Select()
    return problems, [chart]


def verify_panel_grid(t: D.Template, grid: D.PanelGrid, builder, sheet,
                      chart, extra_charts: int = 0) -> list[str]:
    """Prove the grid is one chart on one scale, and that it is derived.

    Three separate claims, and a small multiple is worthless without all of
    them: that the panels really are one object, that they really are on one
    ruler, and that what they draw is computed from the figures above rather
    than typed twice.
    """
    problems: list[str] = []
    # One object for the grid, plus any panel that is deliberately not on the
    # grid's scale - the reference panel is a different measure in a different
    # unit and could not share the ruler even in principle.
    if sheet.ChartObjects().Count != 1 + extra_charts:
        problems.append(
            f"{sheet.ChartObjects().Count} chart objects on the sheet, "
            f"expected {1 + extra_charts} - the whole point of the panel "
            f"engine is that a grid of comparable panels is one chart, and "
            f"sixteen charts drift apart the first time a figure changes")

    axis = chart.Chart.Axes(XL_VALUE)
    if abs(axis.MinimumScale) > 1e-9 or abs(axis.MaximumScale - grid.rows) > 1e-9:
        problems.append(
            f"the band axis is {axis.MinimumScale}..{axis.MaximumScale}, not "
            f"0..{grid.rows} - the bands are the scale, so every panel would "
            f"be drawn on a different one")

    # The input block the engine reads must be formulas, not the values it was
    # handed. Checked on a panel that is drawn, not on the empty cell.
    r = builder.in_r0
    if not str(sheet.Cells(r, builder.in_c0).Formula).startswith("="):
        problems.append(
            "the panel engine's input block holds typed variances - edit a "
            "location above and the pins would not move")
    return problems


def build_sheet(excel, sheet, template: D.Template,
                simple: bool = False) -> tuple[list[str], list]:
    """One template onto one sheet: data zone, then chart zone beside it.

    ``simple`` draws the base-tier version: the same data, fewer tiers, no
    callouts. The Template is reduced only where its *data* is what makes it
    complex - the structure panels - and everything else comes from a smaller
    SheetLayout, so both workbooks are built from one set of numbers.
    """
    name = f"{template.id}{template.variant}"
    if simple:
        template = D.simplify(template)
    if L.is_table(name):
        return build_table_sheet(excel, sheet, template)
    if L.is_structure(name):
        return build_structure_sheet(excel, sheet, template)
    if L.is_line(name):
        return build_line_sheet(excel, sheet, template)
    if L.is_xy(name):
        return build_xy_sheet(excel, sheet, template)
    if L.is_tree(name):
        return build_tree_sheet(excel, sheet, template)
    if L.is_panel(name):
        return build_panel_sheet(excel, sheet, template)
    layout = L.simple_layout_for(name) if simple else L.layout_for(name)
    sheet.Name = layout.template_id

    first, last = write_data(sheet, template, layout)
    excel.Calculate()

    # The chart zone starts where the data zone stops. Derived, not declared -
    # this is the whole point of the exercise.
    chart_left = data_zone_width(sheet, layout) + layout.chart_gap

    shapes = add_chart_title(sheet, template, layout, chart_left)
    built = [add_tier(sheet, template, layout, spec, chart_left, first, last)
             for spec in layout.tiers]
    objects = [obj for obj, _ in built]

    # Re-lock after every tier exists: adding a later chart can trigger a
    # relayout of earlier ones.
    for obj, spec in zip(objects, layout.tiers):
        lock_plot_area(obj.Chart, layout, spec.plot_extent, spec.cross_offset,
                       spec.plot_lead)

    # Only now, because both of these are positioned from the *locked* plot
    # geometry - the rule has to land on the data zero and the caption has to
    # sit on the line it names.
    for (obj, needs_rule), spec in zip(built, layout.tiers):
        if needs_rule:
            shapes += draw_double_rule(sheet, obj, layout, spec,
                                       template.tier(spec.key).reference)
    captions = add_tier_captions(sheet, layout, objects, chart_left)
    shapes += captions

    expected = {f"tier_{spec.key}": spec.cross_offset for spec in layout.tiers}
    problems, geometry = verify(objects, expected, layout, shapes, template)
    problems += check_caption_placement(objects, layout, captions)

    print(f"\n{layout.template_id}: data zone A:"
          f"{L.col_letter(layout.last_column)} = "
          f"{data_zone_width(sheet, layout):.1f}pt, charts from {chart_left:.1f}pt")
    # The shared axis is the one worth printing, and which one that is depends
    # on how the tiers stack.
    box, size, lead, span = (("top", "height", "plot T", "plot H")
                             if layout.horizontal else
                             ("left", "width", "plot L", "plot W"))
    print("  tier geometry as Excel actually laid it out:")
    print(f"    {'chart':<14}{box:>9}{size:>9}{lead:>9}{span:>9}{'series':>8}")
    for name, left, width, inside, insidew, count in geometry:
        print(f"    {name:<14}{left:>9.2f}{width:>9.2f}{inside:>9.2f}"
              f"{insidew:>9.2f}{count:>8}")
    print(f"    (variance tiers carry a deliberate "
          f"{layout.category_offset():.2f}pt offset so their elements sit under "
          f"the measure bars)")

    sheet.Range("A1").Select()
    return problems, objects


def _print_range(sheet, template: D.Template) -> str:
    """The rectangle worth printing, which is not the same on every sheet.

    A chart sheet keeps its source data on the left and its charts to the
    right, and only the charts are the deliverable - the data zone is the input
    the reader edits, and printing it would put a spreadsheet next to the
    picture. So the print area is the union of everything *drawn*: the chart
    objects, and the linked text boxes carrying the title block, which sit
    above them in the same zone.

    A table sheet inverts that. There the table *is* the artefact - it is made
    of cells - and the drawn columns sit on top of it. So the whole used range
    goes, and the union of shapes would have printed four narrow bar columns
    and nothing to read them against.
    """
    name = f"{template.id}{template.variant}"
    drawn = _shape_bounds(sheet)
    if L.is_table(name):
        # From the rendered title down, not from row 1. A table sheet's first
        # seven rows are the typed input block - Entity, Measure, Unit and the
        # rest, as label-and-value pairs - and printing them puts the machinery
        # on the page above the report. The title the reader is meant to see is
        # the formula block below them.
        layout = L.table_layout_for(name)
        used = sheet.UsedRange
        last_row = used.Row + used.Rows.Count - 1
        # The used range is wider than the report. T02A and T04A draw their
        # variance columns as bar charts, and a bar chart needs a favourable
        # and an adverse series to change colour on impact, so each drawn
        # column is fed by two more columns of IF/NA formulas parked to the
        # right of the table. They are deliberately visible - someone tracing a
        # bar back to a number should be able to get there - but they are
        # machinery, and printing them put eight columns of #N/A beside the
        # report. The report is the plan, and nothing after it.
        last_col = len(layout.plan(template))
        if drawn:
            # The panels overhang their cells in both directions. Downwards,
            # one spans the threshold row under the last figure, which the used
            # range stops short of. Rightwards, `_fit_panel_plot` grows the
            # chart object until its *plot* is exactly the column width, so the
            # object itself ends about 4pt past the column - Excel's own chart
            # margin, transparent and empty, but a print area that cuts through
            # it is still cutting through a chart.
            #
            # One column of slack and no more. A blank column separates the
            # report from the feed columns, so absorbing the overhang costs an
            # empty column; taking `drawn` at its word would print the #N/A
            # machinery this branch exists to keep off the page.
            last_row = max(last_row, drawn[2])
            if drawn[3] > last_col:
                last_col = min(drawn[3], last_col + 1)
        return sheet.Range(sheet.Cells(layout.entity_row, 1),
                           sheet.Cells(last_row, last_col)).Address

    if not drawn:                               # nothing drawn; print it all
        return sheet.UsedRange.Address
    top, left, bottom, right = drawn
    return sheet.Range(sheet.Cells(top, left),
                       sheet.Cells(bottom, right)).Address


def _shape_bounds(sheet) -> tuple[int, int, int, int] | None:
    """The cell rectangle every drawn shape on the sheet fits inside.

    Comments are shapes too, and they are anchored to the *input* headers in
    the data zone. Including them dragged the print area's left edge back
    across the whole data block, so every chart sheet printed its own plumbing
    beside the chart.
    """
    top = left = 1 << 30
    bottom = right = 0
    for i in range(1, sheet.Shapes.Count + 1):
        shape = sheet.Shapes(i)
        if shape.Type == MSO_COMMENT:
            continue
        tl, br = shape.TopLeftCell, shape.BottomRightCell
        top, left = min(top, tl.Row), min(left, tl.Column)
        bottom, right = max(bottom, br.Row), max(right, br.Column)
    return None if not bottom else (top, left, bottom, right)


def set_page_setup(sheet, template: D.Template) -> None:
    """One landscape page holding the artefact, and no gridlines.

    Modelled on the panel-charts skill's `_page_setup`, which had the easier
    job: one chart on one sheet, so the print area was that chart's own
    rectangle. Here it has to be computed - see `_print_range`.

    Gridlines are a *window* property in Excel, not a sheet one, so they can
    only be turned off for the sheet that is active at the time. Hence the
    activate; there is no way to set it for a sheet in the background.
    """
    sheet.Activate()
    ps = sheet.PageSetup
    ps.PrintArea = _print_range(sheet, template)
    ps.Orientation = XL_LANDSCAPE
    ps.Zoom = False                    # Zoom must go False before FitToPages
    ps.FitToPagesWide = 1
    ps.FitToPagesTall = 1
    for side in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin"):
        setattr(ps, side, 18)          # a quarter inch, in points
    ps.HeaderMargin = 9
    ps.FooterMargin = 9
    ps.CenterHorizontally = True
    ps.CenterVertically = True
    ps.PrintGridlines = False
    # Screen gridlines are a property of the *window's view of the active
    # sheet*, so there is no way to set them for a sheet in the background -
    # hence the Activate above. Guarded because a workbook driven with
    # Visible=False does not always have a window to ask, and losing the
    # gridline setting is not worth losing the build over.
    try:
        sheet.Parent.Windows(1).DisplayGridlines = False
    except Exception:                                         # noqa: BLE001
        pass


def set_read_me_page_setup(sheet) -> None:
    """The cover sheet on one page too, and portrait because it is prose.

    Every other sheet is landscape because a chart is wider than it is tall.
    The Read me is a single column of wrapped text, which is the opposite
    shape, and forcing it landscape only makes the lines longer to read.
    """
    sheet.Activate()
    ps = sheet.PageSetup
    ps.PrintArea = sheet.UsedRange.Address
    ps.Orientation = XL_PORTRAIT
    ps.Zoom = False
    ps.FitToPagesWide = 1
    ps.FitToPagesTall = 1
    for side in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin"):
        setattr(ps, side, 36)          # half an inch; prose wants a margin
    ps.HeaderMargin = 18
    ps.FooterMargin = 18
    ps.PrintGridlines = False


def verify_page_setup(sheet) -> list[str]:
    """Prove the sheet really will print on one page without gridlines.

    Worth checking rather than assuming: `FitToPagesWide` silently does nothing
    while `Zoom` is still a number, so the assignment can succeed and the
    workbook still print across four pages.
    """
    problems: list[str] = []
    ps = sheet.PageSetup
    if not ps.PrintArea:
        problems.append(f"{sheet.Name} has no print area, so printing it would "
                        f"put the input data on the page beside the charts")
    if ps.Zoom is not False:
        problems.append(f"{sheet.Name} still has Zoom set, which overrides "
                        f"FitToPages - it would print across several sheets")
    if (ps.FitToPagesWide, ps.FitToPagesTall) != (1, 1):
        problems.append(f"{sheet.Name} fits to "
                        f"{ps.FitToPagesWide}x{ps.FitToPagesTall} pages, not 1x1")
    if ps.PrintGridlines:
        problems.append(f"{sheet.Name} would print cell gridlines")
    return problems


def build(templates: list[D.Template], out: Path, keep_open: bool = False,
          export_dir: Path | None = None, simple: bool = False,
          doc: Path | None = None) -> int:
    if win32 is None:
        print("error: pywin32 is not installed", file=sys.stderr)
        return 1

    for template in templates:
        D.check_ties(template)                # nothing is built on a broken tie

    # Nor on a layout registry that has drifted. This one is cheap, needs no
    # Excel, and guards the failure that produced a workbook whose bars had
    # collapsed to a fraction of their length while every other check passed.
    drift = L.check_simple_layouts()
    if drift:
        print("error: layout registries out of step:", file=sys.stderr)
        for problem in drift:
            print(f"  {problem}", file=sys.stderr)
        return 1

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = bool(keep_open)
    excel.DisplayAlerts = False
    wb = None
    problems: list[str] = []
    # Chart geometry is the one thing the document cannot read back out of the
    # saved file, so it is collected here while the charts are open rather
    # than by reopening the workbook afterwards.
    chart_facts: dict[str, list] = {}

    try:
        wb = excel.Workbooks.Add()
        while wb.Sheets.Count > 1:
            wb.Sheets(wb.Sheets.Count).Delete()
        write_read_me(wb.Sheets(1), simple)
        wb.Sheets(1).Activate()
        wb.Windows(1).DisplayGridlines = False
        set_read_me_page_setup(wb.Sheets(1))
        problems.extend(verify_page_setup(wb.Sheets(1)))

        # Sheet order follows the sheet *name*, not the order the templates
        # were built in. The build order is historical - the tranches each
        # template was developed in - and means nothing to a reader opening the
        # workbook looking for C07.
        for template in sorted(templates,
                               key=lambda t: f"{t.id}{t.variant}"):
            sheet = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
            found, objects = build_sheet(excel, sheet, template, simple)
            problems.extend(found)
            set_page_setup(sheet, template)
            problems.extend(verify_page_setup(sheet))
            if doc is not None:
                from ibcs_doc import sheet_chart_facts    # noqa: PLC0415
                chart_facts[sheet.Name] = sheet_chart_facts(sheet)

            if export_dir is not None:
                export_dir.mkdir(parents=True, exist_ok=True)
                # Prefixed, because tier names repeat across templates.
                for obj in objects:
                    name = f"{sheet.Name}_{obj.Name}.png"
                    obj.Chart.Export(str(export_dir / name))
                print(f"  exported tier images to {export_dir}")

        wb.Sheets(1).Activate()

        # SaveAs through the COM server can fail on a mapped drive the server
        # cannot see, so always write locally and copy the file into place.
        with tempfile.TemporaryDirectory() as tmp:
            staged = Path(tmp) / out.name
            wb.SaveAs(str(staged), FileFormat=XL_OPEN_XML_WORKBOOK)
            if not keep_open:
                wb.Close(SaveChanges=False)
                wb = None
            out.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(staged, out)
        print(f"Wrote {out}")

        if doc is not None:
            from ibcs_doc import write_doc                # noqa: PLC0415
            print(f"Wrote {write_doc(out, doc, chart_facts, simple)}")

    except Exception as e:                                    # noqa: BLE001
        print(f"error: Excel build failed: {e}", file=sys.stderr)
        return 1
    finally:
        if not keep_open:
            try:
                if wb is not None:
                    wb.Close(SaveChanges=False)
                excel.Quit()
            except Exception:                                 # noqa: BLE001
                pass

    if problems:
        print("\nVerification problems:", file=sys.stderr)
        for p in problems:
            print(f"  {p}", file=sys.stderr)
        return 1

    print("Verification: tiers aligned.")
    return 0


def main(argv: list[str]) -> int:
    sys.stdout.reconfigure(encoding="utf-8")      # the chart names contain Δ
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", type=Path,
                    default=P.build_dir() / "IBCS-templates.xlsx")
    ap.add_argument("--template", default="C03A",
                    help="comma-separated template ids, one sheet each "
                         f"(built: {', '.join(sorted(L.LAYOUTS))})")
    ap.add_argument("--export-dir", type=Path,
                    default=P.build_dir() / "excel")
    ap.add_argument("--keep-open", action="store_true")
    ap.add_argument("--simple", action="store_true",
                    help="draw the base-tier version of each template")
    ap.add_argument("--doc", type=Path,
                    help="also write the manual-build markdown for this exact "
                         "workbook, read back out of the file just written")
    args = ap.parse_args(argv[1:])

    try:
        templates = [L.template_for(name.strip())
                     for name in args.template.split(",") if name.strip()]
    except KeyError as exc:
        print(f"error: {exc.args[0]}", file=sys.stderr)
        return 1
    return build(templates, args.out, args.keep_open, args.export_dir,
                 args.simple, args.doc)


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
