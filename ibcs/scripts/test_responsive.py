"""Prove the built workbook is live: change a figure, watch the notation follow.

This is the test that matters for the shipped artefact. Someone who opens the
workbook without any of this source code should be able to retype a number and
get a correct chart - not just moved bars, but recalculated variances and, where
a variance changes sign, a bar that changes colour with it. Static per-point
formatting looks identical on the build data and lies the moment anyone edits it,
so it cannot be checked by eye on the original numbers.

The check drives Excel, edits one input cell, recalculates, then reads the
derived cells back and samples the exported chart image for the colour actually
rendered.

Usage:
    python test_responsive.py [workbook.xlsx]
"""

from __future__ import annotations

import shutil
import sys
import tempfile
import time
from collections import Counter
from pathlib import Path

try:
    import win32com.client as win32
except ImportError:
    win32 = None

from PIL import Image

XL_VALUE = 2

import ibcs_data as D
import ibcs_layout as L
import ibcs_excel as X
import ibcs_style as S
import ibcs_paths as P

# The complex workbook carries all seventeen sheets. IBCS-templates.xlsx was
# the development book and is retired; the two named workbooks replaced it.
DEFAULT_BOOK = P.build_dir() / "IBCS Charts - Complex Versions.xlsx"

LAYOUT = L.C03A

# June: actual 193 against a prior year of 182.9, so +10 and green. Drop it to
# 150 and the variance must become roughly -33 and red. Chosen because it flips
# sign, which is the case static formatting gets wrong.
EDIT_ROW = LAYOUT.first_row + 5
NEW_MEASURE = 150.0

# The text half of the same claim: retype the unit and every caption built from
# it must follow. Deliberately not a plausible unit - a wrong one is obvious in
# the printed before/after, where "kEUR" -> "kEUR" would not be.
NEW_UNIT = "kGBP"


def dominant_variance_colour(png: Path) -> str | None:
    """Which variance colour covers the most pixels in an exported tier."""
    img = Image.open(png).convert("RGB")
    counts = Counter(img.getdata())
    wanted = {
        tuple(int(S.VARIANCE["good"].lstrip("#")[i:i + 2], 16) for i in (0, 2, 4)): "good",
        tuple(int(S.VARIANCE["bad"].lstrip("#")[i:i + 2], 16) for i in (0, 2, 4)): "bad",
    }
    best, best_n = None, 0
    for rgb, name in wanted.items():
        if counts.get(rgb, 0) > best_n:
            best, best_n = name, counts[rgb]
    return best


# RPC_E_CALL_REJECTED. Excel is a single-threaded server: while it is busy with
# something of its own - a repaint, a recalculation, writing out a chart image -
# it refuses the incoming call outright rather than queueing it. Nothing is
# wrong with the workbook and nothing is wrong with the call; the same call
# succeeds a moment later.
EXCEL_BUSY = -2147418111


class ExcelBusy(RuntimeError):
    """Excel refused an incoming call because it was busy elsewhere."""


def sheet_of(wb, name: str, tries: int = 40):
    """``sheet_of(wb, name)``, waiting out an Excel that has not finished.

    Every check opens by asking for its sheet, and that is exactly where the
    rejection lands: the check before it has just exported a chart image or
    forced a recalculation, and Excel is still finishing. Because the call is
    refused rather than queued, waiting is the only thing that helps - so ask
    again, for up to ten seconds, and let anything that is not a busy signal
    through untouched.
    """
    for attempt in range(tries):
        try:
            return wb.Sheets(name)
        except Exception as exc:                              # noqa: BLE001
            busy = bool(exc.args) and exc.args[0] == EXCEL_BUSY
            if not busy or attempt == tries - 1:
                raise
            time.sleep(0.25)


def span_on(sheet, label: str = "span") -> float:
    """The span cell a sheet divides by, found by the label beside it.

    Every family that scales its data writes one, and the label is how the
    sheet says which cell it is - safer than recomputing the arithmetic here,
    because if the two ever disagreed it is the sheet that would be right.
    """
    used = sheet.UsedRange
    for r in range(1, used.Row + used.Rows.Count):
        value = sheet.Cells(r, 1).Value
        if isinstance(value, str) and value.strip() == label:
            return float(sheet.Cells(r, 2).Value)
    return 1.0


def category_x(png_width: int, chart_width_pt: float, index: int,
               count: int = LAYOUT.n_categories) -> int:
    """Pixel centre of one category in an exported tier image.

    Derived from the plot geometry the builder locked, not from a fraction of the
    image: the plot area is inset by a left margin and shifted by the deliberate
    category offset, so a naive fraction lands on the neighbouring bar - which is
    exactly the false negative this function exists to avoid.
    """
    scale = png_width / chart_width_pt
    x_pt = (LAYOUT.plot_inside_lead + LAYOUT.category_offset()
            + LAYOUT.plot_inside_span * (index + 0.5) / count)
    return int(x_pt * scale)


def sample_point_colour(png: Path, x_centre: int, half_width: int = 14) -> str:
    """Classify the variance colour covering one vertical slice of the image."""
    img = Image.open(png).convert("RGB")
    w, h = img.size
    px = img.load()
    good = tuple(int(S.VARIANCE["good"].lstrip("#")[i:i + 2], 16) for i in (0, 2, 4))
    bad = tuple(int(S.VARIANCE["bad"].lstrip("#")[i:i + 2], 16) for i in (0, 2, 4))
    tally = Counter()
    for x in range(max(0, x_centre - half_width), min(w, x_centre + half_width)):
        for y in range(h):
            c = px[x, y]
            if c == good:
                tally["good"] += 1
            elif c == bad:
                tally["bad"] += 1
    return tally.most_common(1)[0][0] if tally else "none"


def shape_text(sheet, name: str) -> str:
    """What a text box actually shows.

    Three accessors, because a *linked* text box does not answer the same way as
    one holding its own string, and which of them works depends on the Excel
    build. An empty answer from the first is not evidence the box is empty.
    """
    shape = sheet.Shapes(name)
    for read in (lambda: shape.TextFrame2.TextRange.Text,
                 lambda: shape.TextFrame.Characters().Text,
                 lambda: shape.DrawingObject.Text):
        try:
            text = read()
        except Exception:                                     # noqa: BLE001
            continue
        if text:
            return str(text)
    return ""


def check_text_is_live(excel, sheet, layout: L.SheetLayout) -> list[str]:
    """Retype the unit and prove every caption on the chart follows.

    The rest of this file proves a retyped figure redraws its bar. This is the
    same claim for the words: nothing a chart says may be typed into the chart.
    A caption reading "dPY kEUR" over a sheet whose unit cell now says kGBP is
    not a cosmetic defect - it states the wrong unit, which is the kind of
    error the notation exists to prevent.
    """
    failures: list[str] = []
    unit_cell = sheet.Cells(layout.title_row("unit"), 2)
    original = unit_cell.Value

    watched = [f"title_{layout.template_id}_subject"]
    watched += [f"caption_{spec.key}" for spec in layout.captioned]
    # Only the boxes whose text is built from the unit can be expected to move.
    # A relative-variance caption is a percentage and carries no unit by design,
    # so it is watched precisely to prove it does *not* change.
    expects_unit = {f"title_{layout.template_id}_subject"}
    expects_unit |= {f"caption_{spec.key}" for spec in layout.captioned
                     if "{unit}" in spec.caption}

    before = {name: shape_text(sheet, name) for name in watched}
    unit_cell.Value = NEW_UNIT
    excel.Calculate()
    after = {name: shape_text(sheet, name) for name in watched}

    for name in watched:
        print(f"  {name:<26} {before[name]!r} -> {after[name]!r}")
        if not after[name]:
            failures.append(f"{name}: no text could be read back")
        elif name in expects_unit and NEW_UNIT not in after[name]:
            failures.append(
                f"{name} still reads {after[name]!r} after the unit cell was "
                f"changed to {NEW_UNIT} - its text is not linked to the cell")
        elif name not in expects_unit and after[name] != before[name]:
            failures.append(
                f"{name} changed to {after[name]!r} when only the unit moved; "
                f"a percentage caption carries no unit")

    unit_cell.Value = original
    excel.Calculate()
    return failures


def check_waterfall_is_live(excel, wb) -> list[str]:
    """Retype one line of the statement and prove the whole walk follows.

    This is the claim a waterfall makes that a bar chart does not. Every bar
    below an edited line has to re-float, because each one starts where the last
    one stopped - so a workbook whose bases are numbers rather than formulas
    looks perfect on the build data and misreports every line under the first
    edit. Checked in the cells rather than in the picture: the walk is
    arithmetic, and the arithmetic is what has to move.
    """
    layout, t = L.C12A, D.C12A
    sheet = sheet_of(wb, layout.template_id)
    col, first = layout.col, layout.first_row
    failures: list[str] = []

    edit = 8                       # Personnel expenses, a cost inside a subtotal
    delta = 100.0
    row = first + edit
    labels = [r.label for r in t.rows]

    def read(key):
        return [sheet.Cells(first + i, col[key]).Value for i in range(len(t.rows))]

    before = {k: read(k) for k in ("ac_level", "ac_base", "ac_sub", "var_abs")}
    sheet.Cells(row, col["ac"]).Value = D.C12A_AC[edit] + delta
    excel.Calculate()
    after = {k: read(k) for k in ("ac_level", "ac_base", "ac_sub", "var_abs")}

    # Personnel expenses subtracts, so a bigger cost lowers every level below it.
    for i in range(edit, len(t.rows)):
        moved = after["ac_level"][i] - before["ac_level"][i]
        if abs(moved + delta) > 0.001:
            failures.append(
                f"{labels[i]}: level moved {moved:+.0f} when {labels[edit]} rose "
                f"by {delta:.0f}; every level below a cost line must fall by it")
            break
    for i in range(edit):
        if abs(after["ac_level"][i] - before["ac_level"][i]) > 0.001:
            failures.append(f"{labels[i]}: level moved, but it is above the edit")
            break

    # The bar itself: longer by the edit, and floating lower by it.
    grew = after["ac_sub"][edit] - before["ac_sub"][edit]
    if abs(grew - delta) > 0.001:
        failures.append(f"{labels[edit]}: bar length changed {grew:+.0f}, expected "
                        f"{delta:+.0f}")
    # Row 13, not 12: a result subtotal is drawn from the origin, so its base is
    # zero however the statement moves. The first bar that actually floats below
    # the edit is the line after it.
    floated = after["ac_base"][13] - before["ac_base"][13]
    if abs(floated + delta) > 0.001:
        failures.append(f"{labels[13]}: base moved {floated:+.0f}, so the bar below "
                        f"the edit did not re-float")

    # And the variance follows, on the row and on its subtotal.
    for i in (edit, 11):
        got = after["var_abs"][i] - before["var_abs"][i]
        if abs(got - delta) > 0.001:
            failures.append(f"{labels[i]}: dPY moved {got:+.0f}, expected {delta:+.0f}")

    print(f"  edited {labels[edit]} AC {D.C12A_AC[edit]} -> "
          f"{D.C12A_AC[edit] + delta:.0f}")
    print(f"  {labels[12]:<18} level {before['ac_level'][12]:>6.0f} -> "
          f"{after['ac_level'][12]:>6.0f}")
    print(f"  {labels[13]:<18} base  {before['ac_base'][13]:>6.0f} -> "
          f"{after['ac_base'][13]:>6.0f}")
    print(f"  {labels[19]:<18} level {before['ac_level'][19]:>6.0f} -> "
          f"{after['ac_level'][19]:>6.0f}")
    print(f"  {labels[11]:<18} dPY   {before['var_abs'][11]:>6.0f} -> "
          f"{after['var_abs'][11]:>6.0f}")

    sheet.Cells(row, col["ac"]).Value = D.C12A_AC[edit]
    excel.Calculate()
    return failures


def check_bridge_is_live(excel, wb) -> list[str]:
    """Retype a state and prove the bridge still lands on the actual total.

    C06F's waterfall is a bridge rather than an accumulation: it starts on the
    prior-year total and every step has to carry it to the actual one. That is a
    claim about the whole column, not about one bar, so the check is that the
    closing level still equals the actual total after an edit - which is exactly
    what a reader trusts the chart for.
    """
    layout, t = L.C06F, D.C06F
    sheet = sheet_of(wb, layout.template_id)
    col = layout.col
    failures: list[str] = []

    entries = layout.sheet_entries(t)
    row_of = {("summary", i): layout.first_row + n
              for n, (k, i) in enumerate(entries) if k == "summary"}
    labels = {r.label: row_of[("summary", i)]
              for i, r in enumerate(t.summary_rows)}
    edit_row = layout.first_row + entries.index(("category", 13))   # Illinois

    def level(label):
        return sheet.Cells(labels[label], col["wf_level"]).Value

    def cell(label, key):
        return sheet.Cells(labels[label], col[key]).Value

    before = (level("AC"), cell("AC", "measure"), cell("ΔPY", "var_abs"))
    original = sheet.Cells(edit_row, col["measure"]).Value
    sheet.Cells(edit_row, col["measure"]).Value = original + 100
    excel.Calculate()
    after = (level("AC"), cell("AC", "measure"), cell("ΔPY", "var_abs"))

    print(f"  edited Illinois AC {original:.0f} -> {original + 100:.0f}")
    print(f"  closing level {before[0]:,.1f} -> {after[0]:,.1f}   "
          f"(AC total {after[1]:,.0f})")
    print(f"  total ΔPY     {before[2]:+,.1f} -> {after[2]:+,.1f}")

    # The step grew by 100, so the walk must now land 100 higher.
    if abs((after[0] - before[0]) - 100) > 0.5:
        failures.append(f"the bridge moved {after[0] - before[0]:+.1f} when a "
                        f"state rose by 100; every step below must carry it")
    # And it must still land on the actual total, which is what the bar says.
    if abs(after[0] - after[1]) > 2.5:
        failures.append(f"the bridge closes at {after[0]:,.1f} but the actual "
                        f"total reads {after[1]:,.0f} - it no longer lands")

    sheet.Cells(edit_row, col["measure"]).Value = original
    excel.Calculate()
    return failures


def check_forecast_bridge_is_live(excel, wb) -> list[str]:
    """Retype a forecast month and prove every total it belongs to follows.

    C05X is the densest of the live claims, because one month feeds four places
    at once: its own column, the plan column it is measured against, the closing
    column's expected half, and the bridge that has to keep landing on that
    closing column. A workbook where any one of those lags is worse than a
    static picture, because it looks current.

    September is edited rather than a measured month, so the check also proves
    the split holds: the expected half must move and the measured half must not.
    """
    layout, t = L.C05X, D.C05X
    sheet = sheet_of(wb, layout.template_id)
    col = layout.col
    failures: list[str] = []

    entries = layout.sheet_entries(t)
    row_of = {lbl: layout.first_row + n
              for n, (k, i) in enumerate(entries) if k == "summary"
              for lbl in (t.summary_rows[i].label,)}
    sep = layout.first_row + entries.index(("category", 8))     # September, FC
    delta = 20.0

    def read():
        closing = row_of["2025 AC+FC"]
        return dict(
            plan=sheet.Cells(row_of["2025 PL"], col["measure"]).Value,
            total=sheet.Cells(closing, col["measure"]).Value,
            measured=sheet.Cells(closing, col["bar_ac"]).Value,
            expected=sheet.Cells(closing, col["bar_fc"]).Value,
            landing=sheet.Cells(closing, col["wf_level"]).Value,
        )

    before = read()
    original = sheet.Cells(sep, col["measure"]).Value
    sheet.Cells(sep, col["measure"]).Value = original + delta
    excel.Calculate()
    after = read()

    print(f"  edited Sep {original:.0f} -> {original + delta:.0f} (a forecast month)")
    for key in ("plan", "total", "measured", "expected", "landing"):
        print(f"  {key:<9} {before[key]:>7,.0f} -> {after[key]:>7,.0f}")

    for key, want in (("plan", delta), ("total", delta),
                      ("expected", delta), ("measured", 0.0)):
        moved = after[key] - before[key]
        if abs(moved - want) > 0.001:
            failures.append(
                f"{key} moved {moved:+.0f} when a forecast month rose by "
                f"{delta:.0f}; expected {want:+.0f}")
    # The claim the chart makes: the walk ends on the column it is drawn to.
    if abs(after["landing"] - after["total"]) > 0.001:
        failures.append(f"the bridge closes at {after['landing']:,.0f} but the "
                        f"closing column reads {after['total']:,.0f}")

    sheet.Cells(sep, col["measure"]).Value = original
    excel.Calculate()
    return failures


def _bgr(hex_colour: str) -> int:
    """Excel stores a colour as BGR, which is the reverse of the palette's hex."""
    r, g, b = (int(hex_colour.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4))
    return b << 16 | g << 8 | r


def check_table_is_live(excel, wb) -> list[str]:
    """Retype one country and prove the whole hierarchy follows it, colour included.

    A table makes a different promise from a chart. A chart says "this bar is
    that number"; a table says "these twenty rows are consistent with each
    other", and the way to break that promise is to type a subtotal. So the
    edit here is one country's actual, and the claim is that four derived
    figures on its own row, its continent, and the world all move - and that the
    ΔPL cell changes colour, because the edit pushes it past the threshold the
    footnote prints.

    The colour is read back from Excel's own evaluation of the conditional
    format rather than recomputed here. Recomputing it would only prove that
    our arithmetic agrees with itself; what is in question is whether the
    workbook applies the threshold it advertises.
    """
    layout, t = L.table_layout_for("T01B"), D.T01B
    sheet = sheet_of(wb, layout.template_id)
    failures: list[str] = []

    # Keyed by (tier, series) rather than by tier alone. A measure tier spends
    # three columns, so a tier-keyed lookup keeps only the last of them and
    # "the AC column" silently becomes two columns further right - which is
    # how this check first read the ΔPY% cell and called it an actual.
    cells = layout.columns(t)["values"]
    first_of = {}
    for (key, index), n in cells.items():
        if index == 0:
            first_of[key] = n

    POLAND, EUROPE, WORLD = 4, 8, 19
    ac_col = cells[("m_november", 2)]           # PY, PL, AC - AC is the third

    def cell(row_index, key, offset=0):
        return sheet.Cells(layout.first_row + row_index, first_of[key] + offset)

    def read():
        return dict(
            poland_ac=sheet.Cells(layout.first_row + POLAND, ac_col).Value,
            poland_dpy=cell(POLAND, "dpy_november").Value,
            poland_dpyp=cell(POLAND, "dpyp_november").Value,
            poland_dpl=cell(POLAND, "dpl_november").Value,
            europe_ac=sheet.Cells(layout.first_row + EUROPE, ac_col).Value,
            europe_dpl=cell(EUROPE, "dpl_november").Value,
            world_ac=sheet.Cells(layout.first_row + WORLD, ac_col).Value,
            dpl_red=cell(POLAND, "dpl_november").DisplayFormat.Font.Color
                    == _bgr(S.VARIANCE["bad"]),
        )

    red = _bgr(S.VARIANCE["bad"])
    before = read()
    edit = sheet.Cells(layout.first_row + POLAND, ac_col)
    original = edit.Value
    edit.Value = 60                     # Poland's actual falls from 86
    excel.Calculate()
    after = read()

    print(f"  edited Poland November AC {original:.0f} -> 60")
    print(f"  Poland  ΔPY {before['poland_dpy']:+.0f} -> {after['poland_dpy']:+.0f}"
          f"   ΔPY% {before['poland_dpyp']:+.1f} -> {after['poland_dpyp']:+.1f}"
          f"   ΔPL {before['poland_dpl']:+.0f} -> {after['poland_dpl']:+.0f}"
          f"   red {before['dpl_red']} -> {after['dpl_red']}")
    print(f"  Europe  AC {before['europe_ac']:,.0f} -> {after['europe_ac']:,.0f}"
          f"   ΔPL {before['europe_dpl']:+.0f} -> {after['europe_dpl']:+.0f}")
    print(f"  World   AC {before['world_ac']:,.0f} -> {after['world_ac']:,.0f}")

    moved = original - 60
    for name, got, want in (
            ("Poland ΔPY", after["poland_dpy"], before["poland_dpy"] - moved),
            ("Poland ΔPL", after["poland_dpl"], before["poland_dpl"] - moved),
            ("Europe AC", after["europe_ac"], before["europe_ac"] - moved),
            ("Europe ΔPL", after["europe_dpl"], before["europe_dpl"] - moved),
            ("World AC", after["world_ac"], before["world_ac"] - moved)):
        if abs(got - want) > 0.005:
            failures.append(f"{name} reads {got:,.2f} after the edit, but a "
                            f"country falling by {moved:.0f} makes it {want:,.2f}"
                            f" - it is typed, not derived")
    expected_rel = (60 - 78) / 78 * 100      # Poland's prior year is 78
    if abs(after["poland_dpyp"] - expected_rel) > 0.01:
        failures.append(f"Poland ΔPY% reads {after['poland_dpyp']:+.2f}, "
                        f"expected {expected_rel:+.2f}")
    if before["dpl_red"]:
        failures.append("Poland ΔPL starts red at -5, above the -20 threshold")
    if not after["dpl_red"]:
        failures.append("Poland ΔPL is -31, past the printed -20 threshold, "
                        "and did not turn red - the footnote states a rule the "
                        "sheet does not apply")

    # The threshold is a cell, so moving it has to move the colouring. If this
    # passes while the previous check passes, the two really are one number.
    threshold = sheet.Cells(layout.footnote_row(t), first_of["dpl_november"])
    kept = threshold.Value
    threshold.Value = -40
    excel.Calculate()
    if cell(POLAND, "dpl_november").DisplayFormat.Font.Color == red:
        failures.append("Poland ΔPL stayed red when the threshold was widened "
                        "to -40 - the red is not reading the footnote cell")
    threshold.Value = kept
    edit.Value = original
    excel.Calculate()

    # And the words: both block headers are formulas over the period input, so
    # the month name cannot be stale in one block and current in the other.
    period = sheet.Cells(layout.title_row("period"), 2)
    was = period.Value
    period.Value = "December 2026"
    excel.Calculate()
    headers = [str(sheet.Cells(layout.block_header_row,
                               first_of["m_november"]).Value),
               str(sheet.Cells(layout.block_header_row,
                               first_of["m_ytd_november"]).Value)]
    print(f"  retyped the period to December 2026 - block headers now {headers}")
    if headers != ["December", "_December"]:
        failures.append(f"the block headers read {headers} after the period was "
                        f"retyped; they should be December and _December")
    period.Value = was
    excel.Calculate()
    return failures

def _row_band_colour(png: Path, rows: int, index: int) -> str:
    """What colour the panel draws on one row, read off the exported chart.

    The panel fills its plot area exactly - that is what verify_panels asserts -
    so row *i* of *n* is the *i*th horizontal band of the image. Reading the
    colour there rather than recomputing it is the point: the question is what
    the workbook drew, not what we think it should have drawn.
    """
    image = Image.open(png).convert("RGB")
    width, height = image.size
    top = int(height * index / rows) + 2
    bottom = int(height * (index + 1) / rows) - 2
    counts = Counter()
    for y in range(top, bottom):
        for x in range(0, width, 2):
            r, g, b = image.getpixel((x, y))
            if r > 235 and g > 235 and b > 235:
                continue
            if r > 150 and g < 110 and b < 110:
                counts["bad"] += 1
            elif 90 < r < 210 and g > 130 and b < 110:
                counts["good"] += 1
    return counts.most_common(1)[0][0] if counts else "none"


def check_panel_is_live(excel, wb, tmp: Path) -> list[str]:
    """Retype an actual and prove the drawn variance follows - length and colour.

    T02A states its variances as bars rather than digits, so the claim it makes
    is different in kind from T01B's. A number that fails to recalculate is
    obvious; a bar that is still green after the variance went negative looks
    entirely plausible, and is the failure this template exists to rule out.

    So the check reads the colour back off the exported chart image. Sweden is
    +5 against plan and drawn green; dropping its actual by 16 makes it -11, and
    the bar has to move to the other side of the zero rule and turn red without
    anything being rebuilt.
    """
    layout, t = L.table_layout_for("T02A"), D.T02A
    sheet = sheet_of(wb, layout.template_id)
    failures: list[str] = []

    cells = layout.columns(t)["values"]
    ac_col = cells[("m_november", 1)]          # PL, AC - AC is the second here
    SWEDEN = 5
    panel = sheet.ChartObjects("panel_dpl_november")
    edit = sheet.Cells(layout.first_row + SWEDEN, ac_col)

    def variance():
        return sheet.Cells(layout.first_row + SWEDEN,
                           cells[("dpl_november", 0)]).Value

    before_png = tmp / "t02a_before.png"
    panel.Chart.Export(str(before_png))
    before = (variance(), _row_band_colour(before_png, len(t.rows), SWEDEN))

    original = edit.Value
    edit.Value = original - 16
    excel.Calculate()

    after_png = tmp / "t02a_after.png"
    panel.Chart.Export(str(after_png))
    after = (variance(), _row_band_colour(after_png, len(t.rows), SWEDEN))

    print(f"  edited Sweden November AC {original:.0f} -> {original - 16:.0f}")
    print(f"  ΔPL {before[0]:+.0f} -> {after[0]:+.0f}, "
          f"the drawn bar reads {before[1]} -> {after[1]}")

    if abs(after[0] - (before[0] - 16)) > 0.005:
        failures.append(f"ΔPL reads {after[0]:+.0f} after an actual fell by 16; "
                        f"it should be {before[0] - 16:+.0f}")
    if before[1] != "good":
        failures.append(f"Sweden starts +5 against plan and should be drawn "
                        f"green; the panel reads {before[1]}")
    if after[1] != "bad":
        failures.append(f"Sweden is -11 against plan after the edit and the "
                        f"panel still reads {after[1]} - the colour is frozen, "
                        f"not following the data")

    # The subtotal is drawn too, and it is the one a typed total would break.
    europe = sheet.Cells(layout.first_row + 8, cells[("dpl_november", 0)]).Value
    if abs(europe - (-36 - 16)) > 0.005:
        failures.append(f"Europe's drawn ΔPL is {europe:+.0f}; a country falling "
                        f"by 16 should carry it from -36 to -52")

    edit.Value = original
    excel.Calculate()
    return failures

def check_statement_table_is_live(excel, wb) -> list[str]:
    """Retype one revenue line and prove the whole statement follows, ratio included.

    T03A's claim is different again from the other two tables. T01B's is about a
    hierarchy - a country reaching its continent and the world - and T02A's is
    about a drawn bar changing colour. A statement's is about a *chain*: one line
    has to carry through five results, and the gross margin has to move because
    both the figure above it and the figure below it moved.

    The ratio is the part worth testing. It is the one row on the sheet that is
    neither typed nor a sum, and the one where a plausible-looking static value
    would survive every other check on this page.
    """
    layout, t = L.table_layout_for("T03A"), D.T03A
    sheet = sheet_of(wb, layout.template_id)
    failures: list[str] = []

    cells = layout.columns(t)["values"]
    ac = cells[("m", 2)]                       # PY, PL, AC
    LICENCES, REVENUE, RESULT, MARGIN, GROUP = 0, 4, 12, 13, 20

    def value(row, column=ac):
        return sheet.Cells(layout.first_row + row, column).Value

    def read():
        return dict(revenue=value(REVENUE), result=value(RESULT),
                    margin=value(MARGIN), group=value(GROUP),
                    group_dpy=value(GROUP, cells[("dpy", 0)]),
                    margin_dpy=value(MARGIN, cells[("dpy", 0)]))

    edit = sheet.Cells(layout.first_row + LICENCES, ac)
    original = edit.Value
    before = read()
    edit.Value = original + 100
    excel.Calculate()
    after = read()

    print(f"  edited Licences AC {original:.1f} -> {original + 100:.1f}")
    print(f"  sales revenue {before['revenue']:.1f} -> {after['revenue']:.1f}   "
          f"operating result {before['result']:.1f} -> {after['result']:.1f}")
    print(f"  group result  {before['group']:.1f} -> {after['group']:.1f}   "
          f"ΔPY {before['group_dpy']:+.1f} -> {after['group_dpy']:+.1f}")
    print(f"  gross margin  {before['margin']:.2f}% -> {after['margin']:.2f}%")

    for name, key in (("sales revenue", "revenue"), ("operating result", "result"),
                      ("group result", "group"), ("group ΔPY", "group_dpy")):
        moved = after[key] - before[key]
        if abs(moved - 100) > 0.005:
            failures.append(f"{name} moved {moved:+.2f} when a revenue line rose "
                            f"by 100 - the statement is not carrying it through")

    # The margin has to be recomputed, not merely nudged: both of its terms
    # changed, so the right answer is not the old one plus anything.
    expected = after["result"] / after["revenue"] * 100
    if abs(after["margin"] - expected) > 0.005:
        failures.append(f"gross margin reads {after['margin']:.2f}% where its own "
                        f"two rows give {expected:.2f}% - it is not a formula")
    if abs(after["margin"] - before["margin"]) < 0.01:
        failures.append("gross margin did not move at all; it is a typed number")

    # And its variance is in percentage points, so it must move with it.
    if abs(after["margin_dpy"] - (after["margin"] - value(MARGIN, cells[("m", 0)]))) > 0.005:
        failures.append("the margin's ΔPY is not the difference of the two margins")

    edit.Value = original
    excel.Calculate()
    return failures

def check_statement_panel_is_live(excel, wb, tmp: Path) -> list[str]:
    """Raise a cost and prove its drawn variance changes side *and* colour.

    This is the claim the whole project has been circling: impact is not sign.
    Amortization is 8.5 under plan, which is a cost saved, so its bar is green
    and points the favourable way. Push the actual above plan and the same bar
    must cross the zero rule *and* turn red - not because anything was rebuilt,
    but because the row subtracts and the workbook knows it.

    A frozen colour is the failure this rules out, and it is the one that looks
    entirely plausible on screen: the bar would be in the right place, the right
    length, and the wrong colour, saying a cost overrun was good news.
    """
    layout, t = L.table_layout_for("T04A"), D.T04A
    sheet = sheet_of(wb, layout.template_id)
    failures: list[str] = []

    cells = layout.columns(t)["values"]
    ac = cells[("m", 1)]                       # PL, AC
    AMORTIZATION, EXPENSES, RESULT = 9, 11, 12
    panel = sheet.ChartObjects("panel_dpl")

    def value(row, column=cells[("dpl", 0)]):
        return sheet.Cells(layout.first_row + row, column).Value

    edit = sheet.Cells(layout.first_row + AMORTIZATION, ac)
    original = edit.Value

    before_png = tmp / "t04a_before.png"
    panel.Chart.Export(str(before_png))
    before = (value(AMORTIZATION), value(EXPENSES),
              sheet.Cells(layout.first_row + RESULT, ac).Value,
              _row_band_colour(before_png, len(t.rows), AMORTIZATION))

    edit.Value = original + 18.5               # 47.5 -> 66.0, now over plan
    excel.Calculate()
    after_png = tmp / "t04a_after.png"
    panel.Chart.Export(str(after_png))
    after = (value(AMORTIZATION), value(EXPENSES),
             sheet.Cells(layout.first_row + RESULT, ac).Value,
             _row_band_colour(after_png, len(t.rows), AMORTIZATION))

    print(f"  edited Amortization AC {original:.1f} -> {original + 18.5:.1f}")
    print(f"  its ΔPL {before[0]:+.1f} -> {after[0]:+.1f}, "
          f"drawn {before[3]} -> {after[3]}")
    print(f"  operating expenses {before[1]:+.1f} -> {after[1]:+.1f}, "
          f"operating result {before[2]:.1f} -> {after[2]:.1f}")

    if before[3] != "good":
        failures.append(f"amortization starts 8.5 under plan - a cost saved - "
                        f"and should be drawn green; the panel reads {before[3]}")
    if after[3] != "bad":
        failures.append(f"amortization is now over plan and the panel still "
                        f"reads {after[3]}; a cost overrun drawn green is the "
                        f"error the notation exists to prevent")
    if abs(after[0] - (before[0] + 18.5)) > 0.005:
        failures.append(f"its ΔPL reads {after[0]:+.2f}, not "
                        f"{before[0] + 18.5:+.2f} - the variance is not derived")
    # A cost rising has to push the total cost up and the result down.
    if abs(after[1] - (before[1] + 18.5)) > 0.005:
        failures.append("operating expenses did not follow the line that feeds it")
    if abs(after[2] - (before[2] - 18.5)) > 0.005:
        failures.append(f"operating result went from {before[2]:.1f} to "
                        f"{after[2]:.1f}; a cost rising by 18.5 must take it down "
                        f"by 18.5")

    edit.Value = original
    excel.Calculate()
    return failures

def check_structure_is_live(excel, wb) -> list[str]:
    """Retype a band and prove the total follows it - and the panels stay comparable.

    C01A's claim is not about a bar changing colour; it is about three panels
    remaining readable against each other. So there are two things to hold: the
    period sum over a column has to be the sum of that column, and all three
    panels have to stay on one scale, because the whole layout is an invitation
    to carry 101.9 from the time series into either breakdown of it.

    The second is the one a chart library breaks by default. Left to itself
    Excel scales each chart to its own tallest column, and three panels of very
    different totals then draw them all the same height.
    """
    layout, t = L.structure_layout_for("C01A"), D.C01A
    sheet = sheet_of(wb, layout.template_id)
    failures: list[str] = []

    # The business-area block starts at the first data row; software is its
    # first band and 2026 its fifth column.
    top = layout.first_row
    software, total_row = top + 1, top + 1 + len(t.structure_panels[0].segments)
    plan_column = 2 + 4

    def total(row=total_row, column=plan_column):
        return sheet.Cells(row, column).Value

    before = total()
    edit = sheet.Cells(software, plan_column)
    original = edit.Value
    edit.Value = original + 10
    excel.Calculate()
    after = total()

    print(f"  edited Software 2026 PL {original:.1f} -> {original + 10:.1f}")
    print(f"  the 2026 period sum {before:.1f} -> {after:.1f}")

    if abs(after - (before + 10)) > 0.005:
        failures.append(f"the period sum reads {after:.1f} after a band rose by "
                        f"10; it should be {before + 10:.1f} - it is typed, not "
                        f"summed")

    # The other two panels are breakdowns of a different column and must not move.
    industry_top = top + len(t.structure_panels[0].segments) + 3
    industry_total = industry_top + 1 + len(t.structure_panels[1].segments)
    if abs(sheet.Cells(industry_total, 2).Value - 101.9) > 0.005:
        failures.append("the industry panel moved when a business area changed; "
                        "it is a breakdown of 2025, not of 2026")

    scales = {obj.Name: obj.Chart.Axes(2).MaximumScale
              for obj in sheet.ChartObjects()}
    print(f"  panel maxima after the edit: {scales}")
    if len({round(v, 6) for v in scales.values()}) > 1:
        failures.append(f"the panels are no longer on one scale: {scales} - "
                        f"101.9 would be drawn three different heights")

    # The plan column is plan whatever its figures say: which column is
    # fictitious is structure, and an edit must not repaint it as an actual.
    area = sheet.ChartObjects("panel_area").Chart
    base = area.SeriesCollection(1).Points(5)
    if base.Format.Line.Visible != -1:
        failures.append("the 2026 column lost its outline when a figure "
                        "changed; the scenario is structure, not data")

    edit.Value = original
    excel.Calculate()
    return failures

def check_stacked_bars_are_live(excel, wb) -> list[str]:
    """Retype one country's channel and prove the hierarchy above it follows.

    C02A is the only template in the library that is a structure chart *and* a
    hierarchy, so its claim is both of C01A's and T01B's at once: the band has
    to change, the row total has to follow it, and the two integrated-legend
    rows - Europe and World - have to carry it up. Those two are where a typed
    total would hide, because they are the only rows on the sheet that are not
    drawn as bars, so nothing about the picture would look wrong if they were
    stale.
    """
    layout, t = L.structure_layout_for("C02A"), D.C02A
    sheet = sheet_of(wb, layout.template_id)
    failures: list[str] = []

    top = layout.first_row
    panel = t.structure_panels[0]
    labels = [r.label for r in t.rows]
    retail_row = top + 1 + [s.label for s in panel.segments].index("Retail")
    total_row = top + 1 + len(panel.segments)
    austria = 2 + labels.index("Austria")
    europe = 2 + labels.index("Europe")
    world = 2 + labels.index("World")

    def value(row, column):
        return sheet.Cells(row, column).Value

    def read():
        return (value(retail_row, austria), value(total_row, austria),
                value(retail_row, europe), value(total_row, europe),
                value(retail_row, world), value(total_row, world))

    before = read()
    edit = sheet.Cells(retail_row, austria)
    original = edit.Value
    edit.Value = original + 50
    excel.Calculate()
    after = read()

    print(f"  edited Austria Retail {original:.0f} -> {original + 50:.0f}")
    print(f"  Austria total {before[1]:.0f} -> {after[1]:.0f}   "
          f"Europe retail {before[2]:.0f} -> {after[2]:.0f}   "
          f"Europe total {before[3]:.0f} -> {after[3]:.0f}")
    print(f"  World retail  {before[4]:.0f} -> {after[4]:.0f}   "
          f"World total {before[5]:.0f} -> {after[5]:.0f}")

    for name, i in (("Austria's total", 1), ("Europe's retail", 2),
                    ("Europe's total", 3), ("World's retail", 4),
                    ("World's total", 5)):
        if abs(after[i] - (before[i] + 50)) > 0.005:
            failures.append(f"{name} reads {after[i]:.0f} after a band rose by "
                            f"50; it should be {before[i] + 50:.0f} - the "
                            f"integrated legend is typed, not summed")

    # The axis is fixed at build so the panels stay comparable, and it runs in
    # fractions of the span the sheet divides by - so the top of it, in kCHF,
    # is that fraction times the span. A value can still outgrow it, and saying
    # so beats letting a clipped bar pass for a short one.
    axis = (sheet.ChartObjects("panel_channel").Chart.Axes(2).MaximumScale
            * span_on(sheet))
    if after[1] > axis:
        failures.append(f"Austria now totals {after[1]:.0f} against an axis of "
                        f"{axis:.0f} - the bar is clipped")

    edit.Value = original
    excel.Calculate()
    return failures


def check_cumulative_is_live(excel, wb) -> list[str]:
    """Retype one month and prove every point after it on the line moves.

    A cumulative line makes a promise no other chart in the library makes: each
    point contains all the ones before it. So a single month has to move its own
    point and every later one by the same amount, and the forecast line - which
    cumulates on top of the actuals rather than from zero - has to move too,
    even though nothing about the forecast itself changed.

    That last part is the one worth testing. A forecast cumulative built from
    the forecast months alone would look perfectly normal and would stop telling
    the truth the moment an actual was corrected.
    """
    layout, t = L.line_layout_for("C07C"), D.C07C
    sheet = sheet_of(wb, layout.template_id)
    failures: list[str] = []

    head = layout.first_row
    rows = {name: head + 2 + i
            for i, (name, *_rest) in enumerate(X.LINE_SERIES)}
    MAR, AUG, DEC = 2, 7, 11

    def cell(name, index):
        return sheet.Cells(rows[name], 2 + index)

    def read():
        return (cell("AC cumulative", MAR).Value,
                cell("AC cumulative", AUG).Value,
                cell("FC cumulative", DEC).Value,
                cell("PL cumulative", AUG).Value)

    before = read()
    edit = cell("AC month", MAR)
    original = edit.Value
    edit.Value = original + 60
    excel.Calculate()
    after = read()

    print(f"  edited March actual {original:.0f} -> {original + 60:.0f}")
    print(f"  cumulative actual   March {before[0]:,.0f} -> {after[0]:,.0f}   "
          f"August {before[1]:,.0f} -> {after[1]:,.0f}")
    print(f"  cumulative forecast December {before[2]:,.0f} -> {after[2]:,.0f}   "
          f"(plan unchanged at {after[3]:,.0f})")

    for name, i in (("the actual cumulative at March", 0),
                    ("the actual cumulative at August", 1),
                    ("the forecast cumulative at December", 2)):
        if abs(after[i] - (before[i] + 60)) > 0.005:
            failures.append(f"{name} reads {after[i]:,.0f} after a month rose "
                            f"by 60; it should be {before[i] + 60:,.0f}")
    if abs(after[3] - before[3]) > 0.005:
        failures.append("the plan cumulative moved when an actual was retyped; "
                        "it is a different line and must not")

    # And the variance the message is about follows from the two lines.
    gap = after[1] - after[3]
    if abs(gap - (before[1] - before[3] + 60)) > 0.005:
        failures.append("the August shortfall did not follow the lines it is "
                        "the difference of")

    edit.Value = original
    excel.Calculate()
    return failures


def check_stock_is_live(excel, wb) -> list[str]:
    """Retype one movement and prove every level after it moves with it.

    A stock is the one quantity on this page nobody measures directly: it is
    what the movements leave behind. So the test is a recurrence rather than a
    sum - correcting a single quarter has to shift the level for that quarter
    *and every quarter after it*, right through the forecast and into the plan,
    because a plan built on a stock that has changed is a plan about a different
    warehouse.

    A typed level would pass every visual check and fail this one.
    """
    layout, t = L.line_layout_for("C08H"), D.C08H
    sheet = sheet_of(wb, layout.template_id)
    failures: list[str] = []

    head = layout.first_row
    rows = {"Increase": head + 3, "Decrease": head + 4,
            "Inventory change": head + 5, "Inventory": head + 6}
    Q3_2024, Q4_2025, Q4_2027 = 6, 11, 19

    def cell(name, index):
        return sheet.Cells(rows[name], 2 + index)

    def read():
        return (cell("Inventory change", Q3_2024).Value,
                cell("Inventory", Q3_2024).Value,
                cell("Inventory", Q4_2025).Value,
                cell("Inventory", Q4_2027).Value)

    before = read()
    edit = cell("Decrease", Q3_2024)
    original = edit.Value
    edit.Value = original + 4
    excel.Calculate()
    after = read()

    print(f"  edited the Q3 2024 decrease {original:.0f} -> {original + 4:.0f}")
    print(f"  that quarter's change {before[0]:+.0f} -> {after[0]:+.0f}, "
          f"its level {before[1]:.0f} -> {after[1]:.0f}")
    print(f"  end of 2025 {before[2]:.0f} -> {after[2]:.0f}   "
          f"end of 2027 {before[3]:.0f} -> {after[3]:.0f}")

    if abs(after[0] - (before[0] - 4)) > 0.005:
        failures.append(f"the change reads {after[0]:+.0f}; taking 4 more out "
                        f"should make it {before[0] - 4:+.0f}")
    for name, i in (("that quarter's level", 1), ("the level at end 2025", 2),
                    ("the level at end 2027", 3)):
        if abs(after[i] - (before[i] - 4)) > 0.005:
            failures.append(f"{name} reads {after[i]:.0f} where it should be "
                            f"{before[i] - 4:.0f} - the stock is not "
                            f"accumulating the movements")

    # The forecast and the plan sit downstream of an actual, which is the whole
    # point: a corrected actual has to reach them.
    if abs(after[3] - before[3]) < 0.005:
        failures.append("the 2027 plan did not move when a 2024 actual was "
                        "corrected; the plan is not standing on the actuals")

    edit.Value = original
    excel.Calculate()
    return failures


def _bubble_pixels(png: Path, colour: str, opacity: float = 1.0) -> int:
    """How many pixels the chart paints in one bubble fill.

    Excel exports the blend rather than the fill, so a translucent series has to
    be looked for at the colour it comes out as over paper. Counted with a small
    tolerance because the exporter antialiases every edge, and a bubble this
    size has a great deal of edge.
    """
    want = [round(int(colour.lstrip("#")[i:i + 2], 16) * opacity + 255 * (1 - opacity))
            for i in (0, 2, 4)]
    img = Image.open(png).convert("RGB")
    return sum(count for pixel, count in Counter(img.getdata()).items()
               if all(abs(pixel[i] - want[i]) <= 3 for i in range(3)))


def check_portfolio_is_live(excel, wb, tmp: Path) -> list[str]:
    """Retype one unit's net sales and prove every other bubble resizes.

    C10D is the only sheet in the workbook with no formula on it, and that is a
    fact about portfolio charts rather than a gap: market attractiveness is a
    score, relative market share is a ratio against a competitor who is not on
    the page, and net sales is money. No arithmetic connects the three, so the
    usual claim - that a total follows its components - has nothing to attach
    to.

    What this template promises instead is a *shared ruler*. Every bubble on the
    page is measured against the largest one, so its size means nothing on its
    own and everything in comparison. Double the biggest bubble and every other
    bubble on the chart must lose half its area - including bubbles in other
    series, which is the part that fails if someone ever "tidies" the chart by
    splitting the three scenarios into three charts. That change looks like an
    improvement, leaves every number on the sheet correct, and quietly puts each
    scenario on a ruler of its own.

    So the test is made on the picture, because the sheet cannot show it.
    """
    layout, t = L.xy_layout_for("C10D"), D.C10D
    sheet = sheet_of(wb, layout.template_id)
    chart = sheet.ChartObjects(f"xy_{layout.template_id}")
    failures: list[str] = []

    dark, dark_opacity = S.bubble_fill("AC")
    blue, _ = S.bubble_fill(D.C10D_ACQUISITION)

    # VAB's prior year, 32 mEUR, is the largest bubble on the page and therefore
    # the one that sets the ruler.
    first, _last = layout.block(t, "PY")
    names = [p.entity for p in t.points if p.scenario == "PY"]
    edit = sheet.Cells(first + names.index("VAB"), 5)

    before_png = tmp / "c10d_before.png"
    chart.Chart.Export(str(before_png))
    before = _bubble_pixels(before_png, dark, dark_opacity)

    original = edit.Value
    edit.Value = original * 2
    excel.Calculate()
    after_png = tmp / "c10d_after.png"
    chart.Chart.Export(str(after_png))
    after = _bubble_pixels(after_png, dark, dark_opacity)

    ratio = after / before if before else 0.0
    print(f"  doubled VAB's prior year {original:.0f} -> {original * 2:.0f} mEUR, "
          f"which is the bubble that sets the ruler")
    print(f"  the actual bubbles - a different series - cover {before:,} px "
          f"before and {after:,} px after, a ratio of {ratio:.2f}")

    if before < 10000:
        failures.append(f"only {before} px of actual fill on the chart; the "
                        f"series is not being drawn at all")
    elif not 0.40 <= ratio <= 0.60:
        failures.append(
            f"doubling the largest bubble left the actual series at "
            f"{ratio:.2f} of its area where it should be about half. Either "
            f"the scenarios are not sharing one bubble group, or the sizes are "
            f"not proportional to area")

    edit.Value = original
    excel.Calculate()

    # And the other half of live: a coordinate is read from its cell too. RFA
    # sits at the unattractive end, which is the whole point of the message;
    # move it across the page and the blue must go with it.
    first, _last = layout.block(t, D.C10D_ACQUISITION)
    names = [p.entity for p in t.points if p.scenario == D.C10D_ACQUISITION]
    edit = sheet.Cells(first + names.index("RFA"), 2)

    def blue_centre(png: Path) -> float:
        """Where the acquisition blue sits across the plot, 0 left and 1 right."""
        img = Image.open(png).convert("RGB")
        want = tuple(int(blue.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4))
        xs = [x for x, y, pixel in
              ((x, y, img.getpixel((x, y)))
               for y in range(0, img.height, 3) for x in range(0, img.width, 3))
              if all(abs(pixel[i] - want[i]) <= 3 for i in range(3))]
        return sum(xs) / len(xs) / img.width if xs else -1.0

    moved_png = tmp / "c10d_moved.png"
    original_x = edit.Value
    left = blue_centre(before_png)
    edit.Value = 0.90
    excel.Calculate()
    chart.Chart.Export(str(moved_png))
    right = blue_centre(moved_png)

    print(f"  moved RFA's attractiveness {original_x:.2f} -> 0.90; the blue "
          f"sits at {left:.2f} of the way across before and {right:.2f} after")

    if left < 0 or right < 0:
        failures.append("the acquisition blue is not on the chart at all")
    elif right - left < 0.25:
        failures.append(
            f"RFA's bubble moved from {left:.2f} to {right:.2f} across the "
            f"plot when its attractiveness went from {original_x:.2f} to 0.90; "
            f"the chart is not reading the cell")

    edit.Value = original_x
    excel.Calculate()
    return failures


def check_panels_are_live(excel, wb) -> list[str]:
    """Raise one location above the average and prove its pin crosses over.

    This is T04A's impact test on a grid. A panel of variances split by *sign*
    passes a lazy version of this check without meaning anything - the split
    has to be forced across zero, so the test raises Salzburg, which is 88%
    below the average in 2017, high enough to come out above it.

    Three things have to happen together and none of them is decoration:

    * the location average moves, because it is derived from the fifteen
      panels rather than typed;
    * *every other panel's* variance moves with it, because they are all
      measured against that average - a grid where only the edited panel moved
      would be fifteen charts that happen to share a page;
    * the edited pin changes colour, because it is now above the line it was
      below, and the colour is the only channel saying which way it went.
    """
    layout, t = L.panel_layout_for("C13D"), D.C13D
    sheet = sheet_of(wb, layout.template_id)
    failures: list[str] = []

    head = layout.first_row
    first_panel = head + 2
    keys = [key for key, _label, _values in D.C13D_PANELS]
    rows = {key: first_panel + i for i, key in enumerate(keys)}
    average = first_panel + len(keys)
    column = 2                                   # 2017, the first actual year

    def read(row):
        return sheet.Cells(row, column).Value

    before_avg = read(average)
    before_cologne = read(rows["cologne"])
    edit = sheet.Cells(rows["salzburg"], column)
    original = edit.Value
    # Enough to carry Salzburg from 88% below the average to above it, and to
    # drag the average up under everyone else while it goes.
    edit.Value = original + 900
    excel.Calculate()
    after_avg = read(average)

    print(f"  raised Salzburg 2017 {original:.0f} -> {original + 900:.0f} mEUR")
    print(f"  the location average {before_avg:.1f} -> {after_avg:.1f}")

    if abs(after_avg - (before_avg + 900 / len(keys))) > 0.05:
        failures.append(
            f"the average reads {after_avg:.2f} after one location rose by "
            f"900; it should be {before_avg + 900 / len(keys):.2f} - it is "
            f"typed, not derived from the panels")

    # The panel engine's own input block, which is what the chart reads.
    builder_r0 = None
    for r in range(average + 1, average + 40):
        if str(sheet.Cells(r, 2).Formula).startswith("="):
            builder_r0 = r
            break
    if builder_r0 is None:
        failures.append("no derived variance block under the source data")
        edit.Value = original
        excel.Calculate()
        return failures

    order = [c.key for c in sorted(t.panel_grids["uniform"].cells,
                                   key=lambda c: (c.row, c.col))]
    salzburg_row = builder_r0 + order.index("salzburg")
    cologne_row = builder_r0 + order.index("cologne")
    salzburg = sheet.Cells(salzburg_row, 2).Value
    cologne = sheet.Cells(cologne_row, 2).Value
    print(f"  Salzburg's variance {(original - before_avg) / before_avg * 100:+.0f}%"
          f" -> {salzburg:+.0f}%")
    print(f"  Cologne, untouched, {(before_cologne - before_avg) / before_avg * 100:+.0f}%"
          f" -> {cologne:+.0f}%")

    if salzburg <= 0:
        failures.append(
            f"Salzburg is still {salzburg:+.1f}% after being raised well above "
            f"the average - its pin would stay red and point down")
    if cologne >= 0:
        failures.append(
            f"Cologne reads {cologne:+.1f}% and should have gone negative when "
            f"the average rose under it - the panels are not measured against "
            f"a shared average")

    # One chart for the grid, still one scale - plus the reference panel,
    # which is a separate object on purpose. The fifteen panels are variances
    # in per cent; the reference is the average itself in mEUR. There is no
    # ruler the two can share, so putting it inside the grid would mean two
    # units on one axis, which is the thing the grid exists to prevent.
    names = {sheet.ChartObjects(i).Name
             for i in range(1, sheet.ChartObjects().Count + 1)}
    if names != {"panel_grid", "reference_panel"}:
        failures.append(f"chart objects are {sorted(names)}; expected the grid "
                        f"and the reference panel, and nothing else")
    if "panel_grid" not in names:
        # Nothing further can be measured, and raising here would report a
        # missing object as a crashed suite. The line above already said it.
        edit.Value = original
        excel.Calculate()
        return failures
    axis = sheet.ChartObjects("panel_grid").Chart.Axes(XL_VALUE)
    if abs(axis.MaximumScale - t.panel_grids["uniform"].rows) > 1e-9:
        failures.append(
            f"the band axis ran to {axis.MaximumScale} after the edit, not "
            f"{t.panel_grids['uniform'].rows} - the bands have moved and the "
            f"panels are no longer on one ruler")

    edit.Value = original
    excel.Calculate()
    return failures


def check_tree_is_live(excel, wb) -> list[str]:
    """Retype one base measure and prove all three ratios follow it.

    This is the check the template exists for. A tree draws connectors that
    *state a calculation* - return over net sales is return on sales, and the
    two ratios multiply to ROI - and a sheet whose ratios were typed would keep
    drawing those connectors over numbers that no longer stood in that relation.
    Nothing about the picture would change; it would simply stop being true.

    Net sales is the measure to move, because it is the one feeding two links:
    it is the denominator of return on sales and the numerator of capital
    turnover, so a single edit has to move both middle boxes *and* leave ROI
    alone. ROI is return over capital, and neither of those was touched - so if
    ROI moves, the sheet is computing the tree the wrong way round.

    The scales must also survive the edit. Excel rescales a chart to its own
    data by default, and a box that quietly rescaled would draw the new figure
    at the old height while its neighbours kept the old ruler.
    """
    layout, t = L.tree_layout_for("C11A"), D.C11A
    sheet = sheet_of(wb, layout.template_id)
    failures: list[str] = []

    head = layout.first_row
    rows = {"return": head + 4, "net_sales": head + 5, "capital": head + 6,
            "ros": head + 7, "turnover": head + 8, "roi": head + 9}
    column = 2 + 3                                   # 2024, an actual year

    def read(key):
        return sheet.Cells(rows[key], column).Value

    before = {k: read(k) for k in rows}
    edit = sheet.Cells(rows["net_sales"], column)
    original = edit.Value
    edit.Value = original * 2
    excel.Calculate()
    after = {k: read(k) for k in rows}

    print(f"  doubled net sales 2024 {original:.1f} -> {original * 2:.1f}")
    for key in ("ros", "turnover", "roi"):
        print(f"    {key:<9} {before[key]:8.3f} -> {after[key]:8.3f}")

    # Halving return on sales and doubling capital turnover is what doubling
    # the denominator of one and the numerator of the other has to do.
    if abs(after["ros"] - before["ros"] / 2) > 0.005:
        failures.append(
            f"return on sales reads {after['ros']:.3f} after its denominator "
            f"doubled; it should be {before['ros'] / 2:.3f} - it is typed, not "
            f"derived")
    if abs(after["turnover"] - before["turnover"] * 2) > 0.005:
        failures.append(
            f"capital turnover reads {after['turnover']:.3f} after its "
            f"numerator doubled; it should be {before['turnover'] * 2:.3f} - "
            f"it is typed, not derived")
    # And the product of the two must still be ROI, which net sales cancels out
    # of entirely. This is the whole tree in one line.
    if abs(after["roi"] - before["roi"]) > 0.005:
        failures.append(
            f"ROI moved from {before['roi']:.3f} to {after['roi']:.3f} when net "
            f"sales changed; net sales cancels out of return/capital, so the "
            f"sheet is not computing the tree the way the connectors claim")
    if abs(after["ros"] / 100 * after["turnover"] * 100 - after["roi"]) > 0.005:
        failures.append(
            "return on sales x capital turnover no longer equals ROI after the "
            "edit - the two routes through the tree have parted company")

    # The scales *must* follow the data - that is what the scale block is for -
    # and boxes sharing a unit must follow it together. This check used to
    # assert the opposite, that the rate stayed at its build-time constant,
    # which was true when the tree was fitted once and never again. Now that
    # the boxes divide by a live span, a rate that did not move would mean the
    # sheet was ignoring the edit.
    #
    # What still has to hold, and is the only thing a reader can be misled by,
    # is that the boxes of one unit agree with each other. Doubling net sales
    # takes all three currency boxes from 3.209 to 2.005 points per kEUR - the
    # same number in all three, so a figure carried between them still reads
    # the same size.
    for group, keys in t.tree.groups().items():
        drawn = {}
        # Multiplied back by the group's span, because the axis runs in
        # fractions of it. Without that this measures points per fraction,
        # which is a different number for every box and says nothing.
        span = span_on(sheet, f"span {group}")
        for key in keys:
            chart = sheet.ChartObjects(f"tree_{key}").Chart
            axis = chart.Axes(2)
            drawn[key] = (chart.PlotArea.InsideHeight
                          / ((axis.MaximumScale - axis.MinimumScale) * span))
        print(f"    {group:<9} pt per unit after the edit: "
              + ", ".join(f"{k} {v:.2f}" for k, v in drawn.items()))
        spread = max(drawn.values()) - min(drawn.values())
        if spread / max(drawn.values()) > 0.02:
            detail = ", ".join(f"{k} {v:.3f}" for k, v in drawn.items())
            failures.append(
                f"the {group} boxes parted company after the edit ({detail}) - "
                f"they share a unit, so the same figure would now be two "
                f"different heights")

    edit.Value = original
    excel.Calculate()
    return failures


def check_scattergram_is_live(excel, wb) -> list[str]:
    """Retype one product's margin and prove the sentence over the chart follows.

    C09's message states a count - 45 products of one line above a gross-profit
    threshold - and a count is the most quietly wrong number a report can carry.
    Nothing on the page shows its working: the reader sees a cloud of points and
    a sentence, and has no way to tell whether the sentence was recalculated or
    typed in once and left. So the count is a formula over the products, and
    this is the check that it is.

    The edit is chosen to cross the boundary. VA-19 sits well inside the segment;
    cutting its margin drops it under the 3 mUSD curve, and three things have to
    move together - its own gross profit, its position, and the count. A typed
    count survives all three and goes on saying 45.

    The curves must *not* move. They are a property of the axes, not of the
    data: the locus of 3 mUSD gross profit is where it is whether or not any
    product sits on it, and a curve that shifted when a product moved would mean
    it had been fitted to the points rather than derived.
    """
    layout, t = L.xy_layout_for("C09C"), D.C09C
    sheet = sheet_of(wb, layout.template_id)
    failures: list[str] = []

    first, last = layout.block(t, "VA")
    names = [p.entity for p in t.points if p.group == "VA"]
    row = first + names.index("VA-19")
    count_row = last + 2
    # The curve sample furthest from the axis, where an error would show most.
    curve_row = count_row + 4

    def read():
        return (sheet.Cells(row, 5).Value,
                sheet.Cells(count_row, 5).Value,
                sheet.Cells(curve_row, 5).Value)

    before = read()
    edit = sheet.Cells(row, 2)
    original = edit.Value
    # Enough to put a 32.5 mUSD product under the 3 mUSD curve.
    edit.Value = 5.0
    excel.Calculate()
    after = read()

    print(f"  cut VA-19's margin {original:.2f}% -> 5.00%")
    print(f"  its gross profit {before[0]:.2f} -> {after[0]:.2f} mUSD, which "
          f"takes it out of the {D.C09C_SEGMENT:.0f} mUSD segment")
    print(f"  the count in the segment {before[1]:.0f} -> {after[1]:.0f} "
          f"(the message states {D.C09C_MESSAGE_COUNT})")

    sales = sheet.Cells(row, 3).Value
    want = sales * 5.0 / 100.0
    if after[0] is None or abs(after[0] - want) > 0.005:
        failures.append(f"VA-19's gross profit reads {after[0]} where net sales "
                        f"of {sales:.2f} at 5% is {want:.2f} - it is typed, not "
                        f"the product of the two axes")
    if before[0] < D.C09C_SEGMENT:
        failures.append(f"VA-19 starts at {before[0]:.2f} mUSD, already outside "
                        f"the segment; the edit cannot prove anything")
    if after[1] is None or abs(after[1] - (before[1] - 1)) > 0.005:
        failures.append(
            f"the count reads {after[1]} after a product left the segment; it "
            f"should have fallen from {before[1]:.0f} to {before[1] - 1:.0f}. "
            f"A count that does not move is a count somebody typed")
    if abs(after[2] - before[2]) > 1e-9:
        failures.append(
            f"the 3 mUSD curve moved from {before[2]:.2f} to {after[2]:.2f} "
            f"when a product moved. It is the locus of a gross profit, not a "
            f"fit through the points, and must not depend on them")

    edit.Value = original
    excel.Calculate()
    restored = sheet.Cells(count_row, 5).Value
    if abs(restored - before[1]) > 0.005:
        failures.append(f"the count did not come back to {before[1]:.0f} when "
                        f"the margin was put back")
    return failures


# RPC_E_CALL_REJECTED. Excel is a single-threaded server: while it is busy
# with something of its own - a repaint, a recalculation, its own start-up - it
# refuses the incoming call outright rather than queueing it. Nothing is wrong
# with the workbook when this happens and nothing is wrong with the call; the
# same call succeeds a moment later. It surfaces here rather than in the
# builders because this suite drives seventeen sheets in one session, exporting
# chart images as it goes, which is the busiest thing anything does to Excel.


def _run_once(argv: list[str]) -> int:
    if win32 is None:
        print("error: pywin32 is not installed", file=sys.stderr)
        return 1
    book = Path(argv[1]) if len(argv) > 1 else DEFAULT_BOOK
    if not book.exists():
        print(f"error: {book} not found - build it first", file=sys.stderr)
        return 1

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    # The workbook is seventeen sheets carrying some two hundred chart objects
    # and shapes, and every `Calculate` here repaints all of them. That repaint
    # is what makes Excel busy enough to start refusing calls - the reads that
    # follow arrive while it is still drawing. Nothing on screen is being
    # looked at, so there is nothing to update.
    excel.ScreenUpdating = False
    excel.EnableEvents = False
    wb = None
    failures: list[str] = []

    # Not a TemporaryDirectory context manager: it would try to delete the copy
    # while Excel still had it open, because the cleanup runs before the finally
    # block that quits Excel. Removed after Excel is gone instead.
    tmp = Path(tempfile.mkdtemp())
    try:
        work = tmp / book.name
        work.write_bytes(book.read_bytes())
        wb = excel.Workbooks.Open(str(work))
        # By name, not by index: sheet 1 is the Read me.
        sheet = sheet_of(wb, LAYOUT.template_id)
        chart = sheet.ChartObjects("tier_var_abs").Chart
        ref = LAYOUT.reference(D.C03A)

        def cell(col):
            return sheet.Cells(EDIT_ROW, LAYOUT.col[col])

        def read():
            return {k: cell(k).Value
                    for k in ("measure", "ref", "var_abs", "var_rel")}

        before = read()
        chart_width_pt = chart.ChartArea.Width
        shot_before = tmp / "before.png"
        chart.Export(str(shot_before))
        june_x = category_x(Image.open(shot_before).size[0], chart_width_pt,
                            EDIT_ROW - LAYOUT.first_row)
        colour_before = sample_point_colour(shot_before, june_x)

        print(f"before edit: measure {before['measure']:.0f}, "
              f"{ref} {before['ref']:.1f}, Δ{ref} {before['var_abs']:+.2f}, "
              f"Δ{ref}% {before['var_rel']:+.2f}, June bar reads {colour_before}")

        # The edit. Nothing else is touched - no rebuild, no formatting.
        cell("measure").Value = NEW_MEASURE
        excel.Calculate()

        after = read()
        shot_after = tmp / "after.png"
        chart.Export(str(shot_after))
        colour_after = sample_point_colour(shot_after, june_x)

        print(f"after  edit: measure {after['measure']:.0f}, "
              f"{ref} {after['ref']:.1f}, Δ{ref} {after['var_abs']:+.2f}, "
              f"Δ{ref}% {after['var_rel']:+.2f}, June bar reads {colour_after}")

        expected_abs = NEW_MEASURE - before["ref"]
        expected_rel = expected_abs / before["ref"] * 100

        if abs(after["var_abs"] - expected_abs) > 0.01:
            failures.append(
                f"Δ{ref} did not recalculate: got {after['var_abs']:+.2f}, "
                f"expected {expected_abs:+.2f}")
        if abs(after["var_rel"] - expected_rel) > 0.01:
            failures.append(
                f"Δ{ref}% did not recalculate: got {after['var_rel']:+.2f}, "
                f"expected {expected_rel:+.2f}")
        if abs(after["ref"] - before["ref"]) > 1e-9:
            failures.append(f"{ref} moved; it is a typed input and should not")
        if colour_before != "good":
            failures.append(f"June should start green, read {colour_before}")
        if colour_after != "bad":
            failures.append(
                f"June should turn red once the variance goes negative, "
                f"read {colour_after}")

        print("\nforecast bridge - retyping one month of C05X:")
        failures += check_forecast_bridge_is_live(excel, wb)

        print("\nbridge - retyping one state of C06F:")
        failures += check_bridge_is_live(excel, wb)

        print("\nwaterfall - retyping one line of the statement:")
        failures += check_waterfall_is_live(excel, wb)

        print("\ntable - retyping one country of T01B:")
        failures += check_table_is_live(excel, wb)

        print("\nintegrated panel - retyping one country of T02A:")
        failures += check_panel_is_live(excel, wb, tmp)

        print("\nstatement table - retyping one revenue line of T03A:")
        failures += check_statement_table_is_live(excel, wb)

        print("\nstatement panel - raising a cost line of T04A:")
        failures += check_statement_panel_is_live(excel, wb, tmp)

        print("\nstructure panels - retyping one band of C01A:")
        failures += check_structure_is_live(excel, wb)

        print("\nstacked bars - retyping one channel of C02A:")
        failures += check_stacked_bars_are_live(excel, wb)

        print("\ncumulative lines - retyping one month of C07C:")
        failures += check_cumulative_is_live(excel, wb)

        print("\nstock recurrence - retyping one movement of C08H:")
        failures += check_stock_is_live(excel, wb)

        print("\nportfolio - doubling the bubble that sets the ruler in C10D:")
        failures += check_portfolio_is_live(excel, wb, tmp)

        print("\nscattergram - moving one product across the threshold in C09C:")
        failures += check_scattergram_is_live(excel, wb)

        print("\ndriver tree - doubling net sales in C11A:")
        failures += check_tree_is_live(excel, wb)

        print("\nsmall multiples - raising one location above the average in C13D:")
        failures += check_panels_are_live(excel, wb)

        print("\ntext block - retyping the unit cell:")
        failures += check_text_is_live(excel, sheet, LAYOUT)

        keep = book.parent / "responsive"
        keep.mkdir(parents=True, exist_ok=True)
        for src, name in ((shot_before, "before.png"), (shot_after, "after.png")):
            (keep / name).write_bytes(src.read_bytes())
        print(f"\nchart images kept in {keep}")

        wb.Close(SaveChanges=False)
        wb = None
    except Exception as exc:                                  # noqa: BLE001
        # A busy Excel is not a result. Raised rather than reported so the
        # retry above can tell it apart from a check that actually failed -
        # `finally` still runs, so the retry starts from a clean session.
        if exc.args and exc.args[0] == EXCEL_BUSY:
            raise ExcelBusy(exc) from exc
        print(f"error: {exc}", file=sys.stderr)
        return 1
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
            excel.Quit()
        except Exception:                                     # noqa: BLE001
            pass
        # Only now: Excel holds the copy open until it quits, so removing the
        # directory any earlier fails with a sharing violation - which then
        # masks whatever the real result was.
        shutil.rmtree(tmp, ignore_errors=True)

    if failures:
        print("\nFAILED - the workbook is not fully live:", file=sys.stderr)
        for f in failures:
            print(f"  {f}", file=sys.stderr)
        return 1

    print("\nPASS: variances recalculated, the bar changed colour with the data,"
          " and the chart text followed its cells.")
    return 0


def main(argv: list[str]) -> int:
    """Run the suite, and do not let a busy Excel read as a failed check.

    A retry repeats the output of however far the aborted attempt got, which is
    noisy but honest: it shows what was actually run. Three attempts, because a
    rejection that survives three clean sessions is not a busy Excel any more.
    """
    for attempt in range(1, 4):
        try:
            return _run_once(argv)
        except ExcelBusy as exc:
            print(f"Excel refused a call on attempt {attempt} ({exc}); "
                  f"starting a fresh session", file=sys.stderr)
    print("error: Excel stayed busy across three attempts", file=sys.stderr)
    return 1


if __name__ == "__main__":
    # The scenario deltas print as U+0394 and the Windows console defaults to
    # cp1252, which cannot encode it. Reconfiguring beats making the output
    # ASCII - the notation is the thing being reported on.
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    raise SystemExit(main(sys.argv))
