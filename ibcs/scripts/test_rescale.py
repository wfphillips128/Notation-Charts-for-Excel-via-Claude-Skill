"""Paste bigger numbers over every sheet's inputs and ask whether anything ran off.

This is the check that would have caught the defect the scale block exists to
fix: axis bounds measured for IBCS(R)'s own figures, so somebody else's data
clips or shrinks. It asks two questions of each sheet.

**Did the picture move?** Multiplying an axis bound by the span the sheet
divides by must give back the bound the layout declares. If it does, every
element is drawn at exactly the pixel it was drawn at before, because a chart's
geometry is nothing but its bounds and its values.

**Does it follow the data?** Multiply the typed inputs by a hundred,
recalculate, and read the plotted values and the axis back out of the chart -
back out, rather than recomputed, because what matters is what Excel drew.

Four families, four sets of rules, because they are genuinely different shapes:
a tier stack scales columns, a table scales the feed behind its panels, a
structure sheet and a line sheet scale a mirror of their rows. Three things the
first versions of this got wrong, each of which looked like a defect in the
workbook until it turned out to be a defect here:

* **Only genuinely typed cells move.** A "typed" column on a statement mixes
  line items with subtotals computed from them. Overwriting a subtotal with a
  literal kills the formula *and* re-multiplies a figure its components already
  multiplied - C12A's Sales revenue came out ten thousand times larger.

* **Percentages are not measures.** A typed column in a `rel` group holds a
  ratio. Multiplying it by a hundred models different data, not bigger data,
  and -63% becomes -6300%.

* **A stacked tier is judged on its stack.** Each segment of a waterfall is a
  length; only the running total has to fit inside the axis. Comparing segments
  called a floating bridge clipped when its invisible base held a legitimate
  zero, which is how a row with no bar says so.

Usage:
    python test_rescale.py ["IBCS Charts - Complex Versions.xlsx"]
"""

from __future__ import annotations

import math
import os
import shutil
import sys
import tempfile
import time
from pathlib import Path

try:
    import win32com.client as win32
except ImportError:                                           # noqa: BLE001
    win32 = None

import ibcs_data as D
import ibcs_layout as L
import ibcs_paths as P

XL_VALUE = 2
EXCEL_BUSY = -2147418111
XL_CALC_DONE = 0


def recalc(excel, tries: int = 400) -> None:
    """Recalculate, and wait until Excel has actually finished.

    `Application.Calculate()` returns before the calculation is done on a
    workbook this size - seventeen sheets of charts all repainting - and the
    next COM call is then refused outright with RPC_E_CALL_REJECTED. That
    reads as a flaky test, and it is not: it is this test asking a question
    while Excel is still working. `CalculationState` is the documented way to
    know, and polling it turns an intermittent failure into a short wait.
    """
    excel.Calculate()
    for _ in range(tries):
        try:
            if int(excel.CalculationState) == XL_CALC_DONE:
                return
        except Exception:                                     # noqa: BLE001
            pass                       # busy enough to refuse even this
        time.sleep(0.05)
FACTOR = 100.0

# The complex workbook, which carries all seventeen sheets.
DEFAULT_BOOK = P.build_dir() / "IBCS Charts - Complex Versions.xlsx"


# --------------------------------------------------------------------------- #
# Shared
# --------------------------------------------------------------------------- #
def _typed_cells(sheet, columns, first: int, last: int):
    """Every cell a reader is meant to type over, and nothing else."""
    for column in columns:
        for r in range(first, last + 1):
            cell = sheet.Cells(r, column)
            if str(cell.Formula).startswith("="):
                continue
            if isinstance(cell.Value, (int, float)):
                yield cell


def _span_labelled(sheet, label: str) -> float | None:
    """The span cell a sheet divides by, found by the label written beside it.

    Read rather than recomputed: the sheet is what the charts read, and if the
    two ever disagreed it is the sheet that would be right.
    """
    used = sheet.UsedRange
    for r in range(1, used.Row + used.Rows.Count):
        value = sheet.Cells(r, 1).Value
        if isinstance(value, str) and value.strip() == label.strip():
            return sheet.Cells(r, 2).Value
    return None


def _outside(series, lo: float, hi: float) -> list[float]:
    values = [float(v) for v in (series.Values or ())
              if isinstance(v, (int, float))]
    return [v for v in values if v < lo - 1e-6 or v > hi + 1e-6]


# --------------------------------------------------------------------------- #
# The tier stack: C03A, C04A, C05X, C06F, C12A
# --------------------------------------------------------------------------- #
def _tier_typed(layout) -> list[int]:
    """Typed columns holding a measure rather than a ratio.

    A column's group is declared on its scaled copy, so the scale block is what
    says which of the typed inputs are money and which are per cent.
    """
    groups = {c.scaled_from: c.scale_group for c in layout.columns
              if c.scaled_from}
    return [layout.col[c.key] for c in layout.columns
            if getattr(c, "typed", False) and groups.get(c.key, "unit") != "rel"]


def _tier_plotted(chart, spec, template, lo: float):
    """What a tier draws, as its axis sees it."""
    stacked = spec.stacked or template.tier(spec.key).kind == "waterfall"
    rows, totals = [], {}
    for i in range(1, chart.SeriesCollection().Count + 1):
        series = chart.SeriesCollection(i)
        values = [float(v) for v in (series.Values or ())
                  if isinstance(v, (int, float))]
        if not values:
            continue
        if stacked:
            for j, v in enumerate(values):
                totals[j] = totals.get(j, 0.0) + v
        else:
            rows.append((str(series.Name), min(values), max(values)))
    if stacked and totals:
        heights = list(totals.values())
        # A floating bridge sits well above zero and the rows outside the walk
        # carry a base of nothing and a step of nothing. Their total is exactly
        # zero, which is below the axis floor and is meant to be.
        if lo > 0:
            heights = [h for h in heights if abs(h) > 1e-12]
        if heights:
            rows.append(("the stack", min(heights), max(heights)))
    return rows


def is_simple_book(wb) -> bool:
    """Whether this workbook is the base-tier one, asked of the file itself.

    The Read me says so in as many words, which is a better witness than the
    file name: a copy renamed for review is still the same workbook.
    """
    try:
        sheet = wb.Sheets("Read me")
    except Exception:                                         # noqa: BLE001
        return False
    for row in range(1, 40):
        value = sheet.Cells(row, 1).Value
        if isinstance(value, str) and "This is the simple workbook" in value:
            return True
    return False


def layout_for(name: str, simple: bool):
    """The layout this workbook actually drew the sheet with.

    A simple sheet has its own tiers - fewer of them, with their own geometry -
    so checking it against the full layout asks for charts that are not there.
    That is exactly how a simple-workbook defect survived: the gate could not
    be pointed at the file that had it.
    """
    if simple and L.is_simple(name):
        return L.simple_layout_for(name)
    return L.LAYOUTS[name]


def check_tier_sheets(wb, excel, names, simple: bool = False) -> list[str]:
    failures: list[str] = []
    for name in names:
        sheet = wb.Sheets(name)
        layout, t = layout_for(name, simple), D.TEMPLATES[name]

        for spec in layout.tiers:
            if not spec.scale_group:
                continue                 # clips rather than scales; see below
            span = sheet.Cells(layout.first_row,
                               layout.col[f"span_{spec.scale_group}"]).Value
            axis = sheet.ChartObjects(f"tier_{spec.key}").Chart.Axes(XL_VALUE)
            for edge, got, want in (("bottom", axis.MinimumScale * span,
                                     spec.bounds[0]),
                                    ("top", axis.MaximumScale * span,
                                     spec.bounds[1])):
                if abs(got - want) > 5e-6:
                    failures.append(
                        f"{name}/{spec.key}: the axis {edge} times the span is "
                        f"{got:.6f}, not the {want:g} the layout declares - "
                        f"normalising moved the picture")

        for column in _tier_typed(layout):
            for cell in _typed_cells(sheet, [column], layout.first_row,
                                     layout.last_row()):
                cell.Value = cell.Value * FACTOR
        recalc(excel)

        rates: dict[str, list[tuple[str, float]]] = {}
        for spec in layout.tiers:
            chart = sheet.ChartObjects(f"tier_{spec.key}").Chart
            axis = chart.Axes(XL_VALUE)
            lo, hi = axis.MinimumScale, axis.MaximumScale
            if spec.scale_group:
                rates.setdefault(spec.scale_group, []).append(
                    (spec.key, (hi - lo) / spec.plot_extent))
            for series, low, high in _tier_plotted(chart, spec, t, lo):
                if high > hi + 1e-6 or low < lo - 1e-6:
                    failures.append(
                        f"{name}/{spec.key}: '{series}' plots {low:+.4f}.."
                        f"{high:+.4f} against an axis of {lo:+.4f}..{hi:+.4f} "
                        f"after x{FACTOR:g} - the tier does not follow its data")

        # Tiers sharing a unit must still share a ruler. Same span cell means
        # the same units per point only if the drawn extents agree too.
        for group, entries in rates.items():
            if len(entries) < 2:
                continue
            spread = (max(r for _k, r in entries)
                      - min(r for _k, r in entries))
            if spread > 1e-6:
                detail = ", ".join(f"{k} {r:.8f}/pt" for k, r in entries)
                failures.append(
                    f"{name}: the {group} tiers stopped sharing a scale after "
                    f"x{FACTOR:g} - {detail}")
    return failures


# --------------------------------------------------------------------------- #
# Tables with drawn panels: T02A, T04A
# --------------------------------------------------------------------------- #
def check_table_sheets(wb, excel, names) -> list[str]:
    failures: list[str] = []
    for name in names:
        sheet = wb.Sheets(name)
        layout, t = L.table_layout_for(name), D.TEMPLATES[name]
        if not layout.panels:
            continue                     # a table of numbers draws no bars

        spans, rates = {}, {}
        used = sheet.UsedRange
        for c in range(1, used.Column + used.Columns.Count):
            caption = sheet.Cells(layout.caption_row, c).Value
            if isinstance(caption, str) and caption.startswith("span "):
                spans[caption[5:]] = sheet.Cells(layout.first_row, c).Value

        for key, geometry in layout.panels.items():
            axis = sheet.ChartObjects(f"panel_{key}").Chart.Axes(XL_VALUE)
            span = spans.get(geometry.scale_group, 1.0)
            for edge, got, want in (("bottom", axis.MinimumScale * span,
                                     geometry.bounds[0]),
                                    ("top", axis.MaximumScale * span,
                                     geometry.bounds[1])):
                if abs(got - want) > 5e-6:
                    failures.append(
                        f"{name}/{key}: the axis {edge} times the span is "
                        f"{got:.6f}, not the {want:.6f} its geometry declares")
            rates.setdefault(geometry.scale_group, []).append(
                (key, geometry.width_px
                 / (axis.MaximumScale - axis.MinimumScale) / span))

        for group, entries in rates.items():
            if len(entries) > 1 and (max(r for _k, r in entries)
                                     - min(r for _k, r in entries)) > 1e-4:
                failures.append(
                    f"{name}: the {group} panels are on different rulers - "
                    + ", ".join(f"{k} {r:.4f}px/unit" for k, r in entries))

        plan = layout.plan(t)
        typed = [i + 1 for i, c in enumerate(plan)
                 if getattr(c, "typed", False)]
        for cell in _typed_cells(sheet, typed, layout.first_row,
                                 layout.last_row(t)):
            cell.Value = cell.Value * FACTOR
        recalc(excel)

        for key in layout.panels:
            chart = sheet.ChartObjects(f"panel_{key}").Chart
            axis = chart.Axes(XL_VALUE)
            lo, hi = axis.MinimumScale, axis.MaximumScale
            for i in range(1, chart.SeriesCollection().Count + 1):
                series = chart.SeriesCollection(i)
                out = _outside(series, lo, hi)
                if out:
                    failures.append(
                        f"{name}/{key}: '{series.Name}' plots "
                        f"{max(out, key=abs):+.4f} outside {lo:+.4f}..{hi:+.4f} "
                        f"after x{FACTOR:g}")
    return failures


# --------------------------------------------------------------------------- #
# Structure panels: C01A, C02A
# --------------------------------------------------------------------------- #
def check_structure_sheets(wb, excel, names) -> list[str]:
    failures: list[str] = []
    for name in names:
        sheet = wb.Sheets(name)
        layout, t = L.STRUCTURE_LAYOUTS[name], D.TEMPLATES[name]
        tallest = max(panel.total(i) for panel in t.structure_panels
                      for i in range(len(panel.categories))
                      if not (layout.horizontal and i < len(t.rows)
                              and t.rows[i].kind == "subtotal"))
        maximum = math.ceil(tallest / layout.scale_step) * layout.scale_step
        span = _span_labelled(sheet, "span")

        # The simple workbook draws fewer panels than the template describes,
        # so the panels are taken from the SHEET. Asking for one that is not
        # there used to raise out of the whole run, which is how a workbook
        # this gate could not open kept a defect.
        drawn = {sheet.ChartObjects(i).Name
                 for i in range(1, sheet.ChartObjects().Count + 1)}
        panels = [p for p in t.structure_panels
                  if f"panel_{p.key}" in drawn]
        if len(panels) != len(t.structure_panels):
            print(f"    {name}: {len(panels)} of "
                  f"{len(t.structure_panels)} panels drawn on this sheet")

        tops = [(panel.key,
                 sheet.ChartObjects(f"panel_{panel.key}")
                 .Chart.Axes(XL_VALUE).MaximumScale)
                for panel in panels]
        for key, top in tops:
            if abs(top * span - maximum) > 5e-6:
                failures.append(
                    f"{name}/{key}: the axis top times the span is "
                    f"{top * span:.6f}, not the {maximum:g} the panels share")
        if len({round(top, 12) for _k, top in tops}) != 1:
            failures.append(
                f"{name}: the panels are on different scales - "
                + ", ".join(f"{k} {v:.6f}" for k, v in tops))

        for cell in _typed_cells(sheet,
                                 range(2, 2 + layout.max_categories),
                                 1, sheet.UsedRange.Row
                                 + sheet.UsedRange.Rows.Count - 1):
            cell.Value = cell.Value * FACTOR
        recalc(excel)

        for panel in panels:
            chart = sheet.ChartObjects(f"panel_{panel.key}").Chart
            top = chart.Axes(XL_VALUE).MaximumScale
            for i in range(1, chart.SeriesCollection().Count + 1):
                series = chart.SeriesCollection(i)
                out = _outside(series, float("-inf"), top)
                if out:
                    failures.append(
                        f"{name}/{panel.key}: '{series.Name}' reaches "
                        f"{max(out):.4f} over an axis top of {top:.4f} "
                        f"after x{FACTOR:g}")
    return failures


# --------------------------------------------------------------------------- #
# Line sheets: C07C, C08H
# --------------------------------------------------------------------------- #
def check_line_sheets(wb, excel, names) -> list[str]:
    failures: list[str] = []
    for name in names:
        sheet = wb.Sheets(name)
        layout = L.LINE_LAYOUTS[name]
        span = _span_labelled(sheet, "span")
        charts = [sheet.ChartObjects(i).Name
                  for i in range(1, sheet.ChartObjects().Count + 1)]

        for chart_name in charts:
            axis = sheet.ChartObjects(chart_name).Chart.Axes(XL_VALUE)
            if chart_name == "c08_change":
                # Declares no bounds: Excel fits it, so it already follows its
                # data and there is nothing to divide.
                continue
            got = axis.MaximumScale * span
            if abs(got - layout.maximum) > 5e-6:
                failures.append(
                    f"{name}/{chart_name}: the axis top times the span is "
                    f"{got:.6f}, not the {layout.maximum:g} declared")

        used = sheet.UsedRange
        last = used.Row + used.Rows.Count - 1
        # From column 1, not column 2. C08H's opening balance is a typed input
        # sitting in the *label* column - the stock every level is built from -
        # and leaving it behind scales the flows without the level they move,
        # which is a different chart rather than a bigger one.
        for r in range(1, last + 1):
            if sheet.Rows(r).Hidden:
                continue
            for cell in _typed_cells(sheet, range(1, 14), r, r):
                cell.Value = cell.Value * FACTOR
        recalc(excel)

        for chart_name in charts:
            chart = sheet.ChartObjects(chart_name).Chart
            axis = chart.Axes(XL_VALUE)
            lo, hi = axis.MinimumScale, axis.MaximumScale
            for i in range(1, chart.SeriesCollection().Count + 1):
                series = chart.SeriesCollection(i)
                out = _outside(series, lo, hi)
                if out:
                    failures.append(
                        f"{name}/{chart_name}: '{series.Name}' plots "
                        f"{max(out, key=abs):+.4f} outside {lo:+.4f}..{hi:+.4f} "
                        f"after x{FACTOR:g}")
    return failures


# --------------------------------------------------------------------------- #
def _run_once(book: Path) -> int:
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    # Seventeen sheets of charts repaint on every recalculation, and that
    # repaint is what makes Excel busy enough to start refusing calls.
    excel.ScreenUpdating = False
    excel.EnableEvents = False
    tmp = tempfile.mkdtemp()
    failures: list[str] = []
    try:
        # A unique basename. Excel keys open workbooks by name rather than by
        # path, so a second copy called the same thing hands back the first -
        # and the multiply below then lands on data already multiplied.
        work = os.path.join(tmp, f"rescale_{book.name}")
        shutil.copyfile(book, work)
        wb = excel.Workbooks.Open(work)
        try:
            simple = is_simple_book(wb)
            present = {sheet.Name for sheet in wb.Sheets}
            print(f"  {'simple' if simple else 'complex'} workbook, "
                  f"{len(present) - 1} template sheets")
            for title, fn, names in (
                    ("tier stack", check_tier_sheets,
                     ["C03A", "C04A", "C05X", "C06F", "C12A"]),
                    ("tables", check_table_sheets, ["T02A", "T04A"]),
                    ("structure", check_structure_sheets, ["C01A", "C02A"]),
                    ("lines", check_line_sheets, ["C07C", "C08H"])):
                # Only what this workbook holds. Naming a sheet that is not
                # there must read as "not in this book", never as a pass.
                names = [n for n in names if n in present]
                if not names:
                    print(f"  {title:<12} {'-':<28} not in this workbook")
                    continue
                found = (fn(wb, excel, names, simple)
                         if fn is check_tier_sheets else fn(wb, excel, names))
                print(f"  {title:<12} {', '.join(names):<28} "
                      f"{'clean' if not found else str(len(found)) + ' problem(s)'}")
                failures.extend(found)
        finally:
            wb.Close(SaveChanges=False)
    finally:
        try:
            excel.Quit()
        except Exception:                                     # noqa: BLE001
            pass
        shutil.rmtree(tmp, ignore_errors=True)

    if failures:
        print("\nFAILED - a sheet does not follow its data:", file=sys.stderr)
        for f in failures:
            print(f"  {f}", file=sys.stderr)
        return 1
    print(f"\nPASS: every scaled tier reproduced its declared bounds, and none "
          f"of them clipped when the inputs grew {FACTOR:g}x.")
    print("      C09C and C10D are fitted at build instead, because they are "
          "the only\n      charts that show a value axis - see decisions.md.")
    return 0


def main(argv: list[str]) -> int:
    if win32 is None:
        print("error: pywin32 is not installed", file=sys.stderr)
        return 1
    book = Path(argv[1]) if len(argv) > 1 else DEFAULT_BOOK
    if not book.exists():
        print(f"error: {book} not found - build it first", file=sys.stderr)
        return 1
    for attempt in range(1, 4):
        try:
            return _run_once(book)
        except Exception as exc:                              # noqa: BLE001
            if not (exc.args and exc.args[0] == EXCEL_BUSY) or attempt == 3:
                raise
            print(f"Excel refused a call on attempt {attempt}; "
                  f"starting a fresh session", file=sys.stderr)
    return 1


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    raise SystemExit(main(sys.argv))
