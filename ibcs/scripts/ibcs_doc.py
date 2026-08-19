"""Write the manual-build document for a workbook this skill just produced.

The idea is borrowed from ``panel-charts/scripts/panel_doc.py`` and it is the
only reason to trust a document like this: it opens the ``.xlsx`` that was just
written and **reads the formulas back out of it**. A hand-written guide drifts
the first time a formula changes and nothing catches it, because prose has no
tests. Everything quoted below came out of the file it describes.

What it produces
----------------
One markdown per workbook, in two halves.

* **What every sheet has in common** - the techniques, each illustrated with a
  real formula found by searching the workbook for it, plus the traps that fail
  silently rather than raising.
* **Sheet by sheet** - for each template, the cells you type, the formulas
  behind everything else grouped by their *shape* rather than listed one by
  one, the chart objects and their axis ranges, and the print area.

Grouping by shape is what makes the second half readable. A column of twelve
variance formulas is one idea, not twelve, so every formula is normalised to
R1C1 - ``=D13-C13`` in E13 becomes ``=RC[-1]-RC[-2]`` - and cells that share a
normalised form are reported once, with the range they cover and one worked
example in ordinary A1.

Usage
-----
    python ibcs_excel.py --template C03A,C04A --out book.xlsx --doc book.md

or, against a workbook that already exists::

    python ibcs_doc.py --xlsx book.xlsx --out book.md
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ibcs_layout as L                                       # noqa: E402

try:                                                          # optional
    import win32com.client as win32
except ImportError:                                           # noqa: BLE001
    win32 = None

XL_VALUE = 2

#: A pipe inside a markdown table cell has to be escaped.
BAR = chr(92) + chr(124)

#: Rows of the per-sheet formula table before it stops listing. Whatever is
#: dropped is counted out loud - a table that silently truncates reads as
#: "that is all of them", which is the one thing it must not say.
TABLE_CAP = 40

#: Excel stores functions added after 2007 with an ``_xlfn.`` prefix, which is
#: a storage detail and not what anybody types. Stripped for display, and said
#: out loud once in the document so a reader who opens the XML is not puzzled.
XLFN = re.compile(r"_xlfn\.")

#: One A1 reference, with optional absolute markers and an optional sheet name.
REF = re.compile(r"(?<![A-Za-z0-9_])(\$?)([A-Z]{1,3})(\$?)([1-9][0-9]{0,6})"
                 r"(?![A-Za-z0-9_(])")

#: Formula fragments that identify a technique, in the order they are looked
#: for. Used to find a real example rather than to invent one.
TECHNIQUE_PROBES = {
    "subject": lambda f: f.startswith("=") and '&" in "&' in f,
    "caption": lambda f: f.startswith('="') and "&$B$" in f,
    "variance": lambda f: bool(re.fullmatch(r"=[A-Z]+\d+-[A-Z]+\d+", f)),
    "relative": lambda f: "=IF(" in f and "/" in f and "*100)" in f,
    "split": lambda f: ">0," in f and "NA()" in f and f.count("IF") == 1,
    "span": lambda f: "AGGREGATE(4,6" in f,
    "scaled": lambda f: bool(re.fullmatch(r"=[A-Z]+\d+/\$[A-Z]+\$\d+", f)),
    "text": lambda f: "TEXT(" in f and ("ISNA(" in f or "ISBLANK(" in f),
    "cascade": lambda f: bool(re.fullmatch(r"=[A-Z]+\d+\+[A-Z]+\d+\*[A-Z]+\d+",
                                           f)),
    "subtotal": lambda f: f.startswith("=SUM("),
    "clip": lambda f: "MEDIAN(" in f,
    "marker": lambda f: "UNICHAR(" in f,
}


# --------------------------------------------------------------- formatting
def clean(formula: str) -> str:
    """A formula as the user would type it."""
    return XLFN.sub("", formula or "")


def col_letters(col: int) -> str:
    s, c = "", col
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def addr(row: int, col: int) -> str:
    return "%s%d" % (col_letters(col), row)


def r1c1(formula: str, row: int, col: int) -> str:
    """The formula with every reference made relative to its own cell.

    Two cells hold "the same formula" when they say the same thing about their
    own position, which is what this measures. It is the whole basis of the
    grouping - without it a twelve-row column reports as twelve findings.
    """
    def one(m):
        cabs, letters, rabs, digits = m.groups()
        c = 0
        for ch in letters:
            c = c * 26 + (ord(ch) - 64)
        r = int(digits)
        cpart = "C%d" % c if cabs else "C[%d]" % (c - col)
        rpart = "R%d" % r if rabs else "R[%d]" % (r - row)
        return rpart + cpart
    return REF.sub(one, formula)


def ranges(cells: list[tuple[int, int]]) -> str:
    """Summarise a set of (row, col) as an A1 range, honestly.

    A full rectangle is named as one. Anything else says how many cells it is
    and where they live, rather than pretending to a shape it does not have -
    a range that quietly widens is exactly the sort of claim this file exists
    to avoid making.
    """
    rows = [r for r, _ in cells]
    cols = [c for _, c in cells]
    r0, r1, c0, c1 = min(rows), max(rows), min(cols), max(cols)
    box = (r1 - r0 + 1) * (c1 - c0 + 1)
    span = "%s:%s" % (addr(r0, c0), addr(r1, c1)) if box > 1 else addr(r0, c0)
    if len(cells) == box:
        return "`%s`" % span
    return "%d cells in `%s`" % (len(cells), span)


# ------------------------------------------------------------ reading a sheet
class SheetFacts:
    """Everything one worksheet says about itself, read rather than assumed."""

    def __init__(self, ws):
        self.ws = ws
        self.name = ws.title
        self.formulas: dict[str, list[tuple[int, int, str]]] = {}
        self.typed: list[tuple[int, int]] = []
        self.labels: list[tuple[int, int]] = []
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if v is None:
                    continue
                if isinstance(v, str) and v.startswith("="):
                    shape = r1c1(clean(v), cell.row, cell.column)
                    self.formulas.setdefault(shape, []).append(
                        (cell.row, cell.column, clean(v)))
                elif isinstance(v, (int, float)):
                    self.typed.append((cell.row, cell.column))
                else:
                    self.labels.append((cell.row, cell.column))
        self.hidden = sorted(d for d, dim in ws.column_dimensions.items()
                             if dim.hidden)

    @property
    def n_formulas(self) -> int:
        return sum(len(v) for v in self.formulas.values())

    def groups(self):
        """Formula groups, largest first, then by where they start."""
        out = []
        for shape, cells in self.formulas.items():
            cells.sort()
            out.append((shape, cells))
        out.sort(key=lambda g: (-len(g[1]), g[1][0]))
        return out

    def header_for(self, row: int, col: int) -> str:
        """What a cell is called: its column header, or failing that its row.

        A header sits directly over its column, so only one row up is looked
        at - two rows up is a caption or a title, not a name for this data.
        Length matters as well as position: the message line sits directly
        over the subject formula and is a paragraph, not a header, so anything
        long is rejected and the row label in column A is used instead.
        """
        above = self.ws.cell(row=max(1, row - 1), column=col).value
        if (isinstance(above, str) and above and not above.startswith("=")
                and len(above) <= 40):
            return above
        left = self.ws.cell(row=row, column=1).value
        if isinstance(left, str) and left and not left.startswith("="):
            return left[:40]
        return ""

    def find(self, test) -> tuple[int, int, str] | None:
        """The first cell whose formula satisfies `test`, in reading order."""
        best = None
        for cells in self.formulas.values():
            for row, col, f in cells:
                if test(f) and (best is None or (row, col) < best[:2]):
                    best = (row, col, f)
        return best


#: The order technique examples are looked for in. Reading order through the
#: workbook would take most of them off C01A, whose formulas carry an extra
#: condition each and are the worst place to meet an idea for the first time.
#: These four are the plainest statement of each technique in the library.
TEACHING_ORDER = ("C03A", "C04A", "C12A", "T04A")


def find_example(sheets: list[SheetFacts], key: str):
    """One real example of a technique: (sheet name, address, formula)."""
    test = TECHNIQUE_PROBES[key]
    rank = {n: i for i, n in enumerate(TEACHING_ORDER)}
    sheets = sorted(sheets, key=lambda x: rank.get(x.name, len(rank)))
    for s in sheets:
        hit = s.find(test)
        if hit:
            return s.name, addr(hit[0], hit[1]), hit[2]
    return None


def quote(example, note: str = "") -> list[str]:
    """A found formula, rendered with the address it was found at."""
    if not example:
        return ["*(no example of this in this workbook)*", ""]
    sheet, where, formula = example
    out = ["```", formula, "```", "",
           "Read from `%s!%s`.%s" % (sheet, where,
                                     (" " + note) if note else ""), ""]
    return out


# ------------------------------------------------------------- chart facts
def chart_facts_from(xlsx_path: str) -> dict:
    """Chart names, geometry and axis ranges, read out of the saved file.

    Only used when the caller did not collect them during the build. Needs
    Excel; without it the document simply omits the chart tables rather than
    guessing at them.
    """
    if win32 is None:
        return {}
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    out = {}
    wb = None
    try:
        wb = excel.Workbooks.Open(os.path.abspath(xlsx_path), ReadOnly=True)
        for sheet in wb.Sheets:
            out[sheet.Name] = sheet_chart_facts(sheet)
    except Exception:                                         # noqa: BLE001
        return {}
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
            excel.Quit()
        except Exception:                                     # noqa: BLE001
            pass
    return out


def sheet_chart_facts(sheet) -> list[dict]:
    """One sheet's chart objects, as plain data. Safe to call mid-build."""
    facts = []
    for obj in sheet.ChartObjects():
        d = {"name": obj.Name,
             "left": round(float(obj.Left), 1), "top": round(float(obj.Top), 1),
             "width": round(float(obj.Width), 1),
             "height": round(float(obj.Height), 1)}
        try:
            chart = obj.Chart
            d["series"] = int(chart.SeriesCollection().Count)
        except Exception:                                     # noqa: BLE001
            pass
        try:
            axis = obj.Chart.Axes(XL_VALUE)
            d["min"] = round(float(axis.MinimumScale), 4)
            d["max"] = round(float(axis.MaximumScale), 4)
        except Exception:                                     # noqa: BLE001
            pass
        facts.append(d)
    return facts


# ----------------------------------------------------------------- the doc
def family_of(template_id: str) -> str:
    if L.is_table(template_id):
        return "table"
    if L.is_structure(template_id):
        return "structure"
    if L.is_line(template_id):
        return "line"
    if L.is_xy(template_id):
        return "XY"
    if L.is_tree(template_id):
        return "tree"
    if L.is_panel(template_id):
        return "panel"
    return "tier stack"


FAMILY_NOTE = {
    "tier stack": "One chart per tier, stacked along the page and sharing a "
                  "category axis. The tiers are separate chart objects "
                  "positioned to line up, because one chart cannot carry two "
                  "value axes that both start at zero.",
    "table": "A printed grid, with some columns *drawn* instead of printed - "
             "the drawn ones are ordinary variance charts sized to the row "
             "pitch and positioned against the rows they belong to.",
    "structure": "Categories stacked inside a column or bar, on one scale "
                 "shared by every panel on the sheet.",
    "line": "Series laid out in rows rather than columns, one row per series "
            "across the periods, with a combo chart over them.",
    "XY": "A scatter or bubble plot. These two are the only charts in the "
          "library that show a value axis, which is why they are fitted at "
          "build time rather than normalised - see the scale block below.",
    "tree": "Boxes and connectors: several small charts and the arithmetic "
            "between them, with boxes that share a unit sharing a scale.",
    "panel": "A grid of small charts drawn as one native chart object by the "
             "`panel-charts` skill, plus a reference panel drawn beside it.",
}


def build_markdown(xlsx_path: str, chart_facts: dict | None = None,
                   simple: bool = False) -> str:
    wb = openpyxl.load_workbook(xlsx_path)
    sheets = [SheetFacts(wb[n]) for n in wb.sheetnames]
    by_name = {s.name: s for s in sheets}
    templates = [s for s in sheets if s.name != "Read me"]
    charts = chart_facts if chart_facts is not None else {}

    out: list[str] = []
    w = out.append
    book = os.path.basename(xlsx_path)

    # ------------------------------------------------------------- heading
    w("# %s - how it is built" % os.path.splitext(book)[0])
    w("")
    w("Every formula quoted here was **read back out of `%s`** after it was "
      "written, so this document and that workbook cannot disagree. Where it "
      "says a formula sits in `C03A!E13`, it does." % book)
    w("")
    if simple:
        w("This is the *simple* workbook: each template drawn with its base "
          "tier only. The complex workbook is the same construction with every "
          "tier present, and the techniques below are identical.")
        w("")
    w("| | |")
    w("|---|---|")
    w("| Sheets | %d (a Read me and %d templates) |"
      % (len(sheets), len(templates)))
    w("| Formula cells | %d |" % sum(s.n_formulas for s in templates))
    w("| Typed cells | %d |" % sum(len(s.typed) for s in templates))
    if charts:
        w("| Chart objects | %d |"
          % sum(len(charts.get(s.name, [])) for s in templates))
    w("| Distinct formula shapes | %d |"
      % len({sh for s in templates for sh in s.formulas}))
    w("")
    w("> **One note on what you will see if you open the XML.** Excel stores "
      "functions added after 2007 with an `_xlfn.` prefix - `_xlfn.AGGREGATE`, "
      "`_xlfn.UNICHAR`. That is a storage detail; you type them without it. "
      "This document strips the prefix everywhere.")
    w("")
    w("---")
    w("")

    # ------------------------------------------------- shared techniques
    w("## What every sheet has in common")
    w("")
    w("Ten techniques carry the whole workbook. Read these once and the "
      "sheet-by-sheet section below is mostly addresses.")
    w("")

    w("### 1. Six typed lines, and every word derived from them")
    w("")
    w("Rows 1 to 6 of every sheet are the only text anybody types: entity, "
      "measure, unit, period, reference scenario, message. The subject line "
      "the reader sees is a formula over them:")
    w("")
    out.extend(quote(find_example(templates, "subject"),
                     "Retype the unit and the subject line follows."))
    w("Tier captions are formulas over the same cells, so a caption cannot "
      "describe a comparison the sheet is not making:")
    w("")
    out.extend(quote(find_example(templates, "caption")))
    w("On a chart sheet the caption *cell* is then linked to a text box on the "
      "chart - two hops, one source. **Nothing a chart says is typed into the "
      "chart.** Type it in a cell and link the box to the cell, every time.")
    w("")

    w("### 2. Two zones, and nothing straddles them")
    w("")
    w("Each sheet is a data zone of declared width on the left and the charts "
      "to its right, with the chart zone's left edge measured from where the "
      "data ends. The print area is the chart zone alone:")
    w("")
    w("| Sheet | Cells used | Print area |")
    w("|---|---|---|")
    for t in templates:
        w("| %s | `%s` | `%s` |"
          % (t.name, t.ws.dimensions, (t.ws.print_area or "-")))
    w("")
    # Which way a sheet's deliverable sits is a per-family fact, not a
    # per-sheet accident, so it is derived from the family rather than
    # from where the print area happens to start.
    tables = [t.name for t in templates if family_of(t.name) == 'table']
    panels = [t.name for t in templates if family_of(t.name) == 'panel']
    if tables:
        w("**A table sheet is the other way round.** On %s the grid *is* the deliverable, so the print area is the grid - starting below the six "
          "typed title rows, which are excluded for exactly the reason a chart sheet excludes its data zone: they are the input." % ", ".join(tables))
        w("")
    if panels:
        w("%s park%s the charts **below** the data rather than beside it, because the panel engine owns its own columns to the right. Same rule, "
          "read the other way: the print area holds the deliverable and nothing else, so it goes wherever the data does not."
          % (", ".join(panels), "" if len(panels) > 1 else "s"))
        w("")
    w("**Never `AutoFit` a data block a chart is positioned against.** One "
      "long string moves the boundary and the charts end up drawn over the "
      "source data. Column widths here are declared, never computed from "
      "content.")
    w("")

    w("### 3. Splitting a series by what it means, not by its sign")
    w("")
    w("A variance is drawn in two colours, and Excel gives one colour to a "
      "series. So the split happens in the *cells*: one column per direction, "
      "each `NA()` where the other owns the point.")
    w("")
    out.extend(quote(find_example(templates, "split")))
    w("`NA()` and not `\"\"` - an empty string plots as a **zero**, which puts "
      "a bar of nothing on the axis where there should be no bar at all. That "
      "one substitution is the most common way to get a chart that looks "
      "broken for no visible reason.")
    w("")
    w("The same split does the work for impact: within one variance column of "
      "a statement the favourable values sit on **both** sides of the axis, "
      "because a cost line up is adverse and a cost line down is favourable. "
      "Split by impact, never by sign, or every cost overrun comes out green.")
    w("")

    w("### 4. The scale block - why the charts follow your numbers")
    w("")
    w("This is the one technique with no equivalent in a hand-drawn chart, and "
      "it exists because **Excel will not bind an axis bound to a formula**. "
      "An axis maximum is a number you type; it cannot be `=MAX(...)`. So a "
      "chart built for one set of figures clips or shrinks the moment somebody "
      "pastes their own over the top.")
    w("")
    w("The way round it is to invert the problem: instead of scaling the axis "
      "to the data, **scale the data to a fixed axis**. Three parts:")
    w("")
    w("**A span cell** - one per group of tiers sharing a unit - holding the "
      "largest magnitude anywhere in that group:")
    w("")
    out.extend(quote(find_example(templates, "span"),
                     "`AGGREGATE(4,6,...)` is MAX ignoring errors, which "
                     "matters because half these columns are deliberately "
                     "`NA()`."))
    w("**A scaled copy of every drawn column**, which is what the chart "
      "actually reads:")
    w("")
    out.extend(quote(find_example(templates, "scaled")))
    w("**Axis bounds divided by the same span.** Dividing both the values and "
      "the bounds by one number is visually a no-op at today's figures and "
      "follows the data at any others. Zero divided by anything is zero, so "
      "the zero line, the reference rules and every caption positioned from "
      "plot geometry stay exactly where they were.")
    w("")
    w("Two consequences worth knowing before you copy this:")
    w("")
    w("- **The scaled columns are hidden, not parked far right.** A hidden "
      "column measures zero width, so the machinery can be added to a sheet "
      "without moving a single chart on it.")
    w("- **It cannot be used where an axis is read.** Every chart in this "
      "library hides its value axis except the XY pair, which show gridlines "
      "and tick labels - and dividing those by a span would print `0.25` where "
      "the data says `29.16%`. Those two are fitted at build time instead: "
      "correct for whatever data is present when the sheet is built, but not "
      "live on edit. That is a stated limit, not an oversight.")
    w("")

    w("### 5. Text columns, because a linked label ignores its own format")
    w("")
    w("A data label linked to a cell shows the **cell's** value and ignores "
      "the label's number format, so a variance of `31.61057692` prints every "
      "one of those digits. The formatting therefore has to happen in a cell:")
    w("")
    out.extend(quote(find_example(templates, "text")))
    w("The label is then linked to that column. The cell says exactly what the "
      "chart says, which is also why the two can be checked against each other.")
    w("")

    w("### 6. Variances, and the relative variance that has no answer")
    w("")
    out.extend(quote(find_example(templates, "variance")))
    w("A relative variance divides by the reference, which may be zero, and "
      "`NA()` is the honest answer - a bar of nothing, not a bar of zero:")
    w("")
    out.extend(quote(find_example(templates, "relative")))

    w("### 7. The waterfall cascade")
    w("")
    w("A waterfall is not a chart type here; it is two columns of arithmetic "
      "and an invisible series. Each row carries a **running level** and the "
      "level it **starts from**, and the sign column says whether the row adds "
      "or subtracts:")
    w("")
    out.extend(quote(find_example(templates, "cascade"),
                     "The running level: the previous level plus this row's "
                     "signed value."))
    w("The floating bar is then the difference between the two, drawn over an "
      "invisible series holding the start. A subtotal row breaks the chain by "
      "reading the level directly instead of adding to it - which is what "
      "makes a subtotal a column from zero rather than a floating step.")
    w("")

    w("### 8. A total is data")
    w("")
    w("A total with components is a formula, never an input. A total that does "
      "not follow its parts is the same lie a static variance colour is:")
    w("")
    out.extend(quote(find_example(templates, "subtotal")))

    w("### 9. Scenario fills, the forecast hatch, and pins")
    w("")
    w("These are formatting rather than formulas, so they are not readable out "
      "of the cells - but they are the notation, so they belong here:")
    w("")
    w("| What | How it is drawn |")
    w("|---|---|")
    w("| AC - actual | solid dark fill |")
    w("| PY - prior year | solid light fill |")
    w("| PL / BU - plan or budget | outlined, no fill |")
    w("| FC - forecast | hatched, 45 degrees ascending |")
    w("| absolute variance | a column or bar, red or green **by impact** |")
    w("| relative variance | a **pin**: a thin stem with a head marker |")
    w("")
    w("A pin has no chart type. It is a line series with the line hidden, "
      "markers on, and a **custom Y error bar** as the stem - the error bar's "
      "weight is in points, so the stem is exactly as thin as it should be, "
      "where a very narrow column bottoms out around 9px because `GapWidth` "
      "caps at 500. The head carries the minuend's scenario fill, which is "
      "what says whether a relative variance was measured or forecast.")
    w("")
    w("Where a variance runs off the end of its panel, the bar is **clipped** "
      "and the label carries the overflow, because a bar drawn at full length "
      "off the plot says nothing about having been cut:")
    w("")
    clip = find_example(templates, "clip") or find_example(templates, "marker")
    out.extend(quote(clip))

    w("### 10. Page setup")
    w("")
    w("Every sheet prints on one page: print area set to the chart zone alone, "
      "`FitToPagesWide/Tall = 1`, worksheet gridlines off. The data zone is "
      "deliberately outside the print area - it is the input, not the "
      "deliverable.")
    w("")

    w("### The traps that fail silently")
    w("")
    w("Each of these produces a wrong-looking sheet rather than an error.")
    w("")
    w("| Trap | What you see |")
    w("|---|---|")
    w("| `\"\"` where `NA()` was meant | a bar of nothing sitting on the axis |")
    w("| `OR()` guarding a lookup | `#REF!` across every spacer row - `OR` "
      "evaluates all its arguments, nested `IF`s short-circuit |")
    w("| a data label typed rather than linked | it stops tracking the cell "
      "the first time a figure changes |")
    w("| a label linked to a raw cell | full floating-point precision printed "
      "beside a rounded chart |")
    w("| `AutoFit` on the data block | charts drawn on top of the source data |")
    w("| an axis bound typed as a number | the chart clips the day somebody "
      "pastes their own figures |")
    w("| variance colour taken from the sign | every cost overrun green |")
    w("")
    w("---")
    w("")

    # ------------------------------------------------------ sheet by sheet
    w("## Sheet by sheet")
    w("")
    w("| Sheet | Family | Formulas | Typed | Charts |")
    w("|---|---|---|---|---|")
    for s in templates:
        w("| [%s](#%s) | %s | %d | %d | %s |"
          % (s.name, s.name.lower(), family_of(s.name), s.n_formulas,
             len(s.typed),
             len(charts.get(s.name, [])) if charts else "-"))
    w("")

    for s in templates:
        out.extend(sheet_section(s, charts.get(s.name, []), by_name))

    w("---")
    w("")
    w("*Generated by `ibcs_doc.py` from the workbook itself.*")
    return "\n".join(out) + "\n"


def sheet_section(s: SheetFacts, charts: list[dict], by_name) -> list[str]:
    out: list[str] = []
    w = out.append
    fam = family_of(s.name)

    w("### %s" % s.name)
    w("")
    message = s.ws.cell(row=6, column=2).value
    if isinstance(message, str) and message:
        w("> %s" % message)
        w("")
    w("**%s.** %s" % (fam.capitalize(), FAMILY_NOTE[fam]))
    w("")
    w("| | |")
    w("|---|---|")
    w("| Cells used | `%s` |" % s.ws.dimensions)
    w("| Print area | `%s` |" % (s.ws.print_area or "-"))
    if s.hidden:
        w("| Hidden engine columns | %s |"
          % ", ".join("`%s`" % h for h in s.hidden))
    w("| Formula cells | %d, in %d shapes |" % (s.n_formulas, len(s.formulas)))
    w("| Typed number cells | %d |" % len(s.typed))
    w("")

    if s.typed:
        w("**Typed values** - %s. Everything else on the sheet is a formula "
          "over these. The cells a reader is meant to edit are shaded; a "
          "sheet may also park a constant the engine needs in this range, so "
          "go by the shading rather than by the count." % ranges(s.typed))
        w("")

    # A group earns a row if it repeats, or if the sheet gave it a name -
    # which is how the span cells, one formula each and the most
    # interesting cells on the sheet, avoid being filed under "one-off".
    groups = s.groups()
    named = [(g, s.header_for(g[1][0][0], g[1][0][1])) for g in groups]
    shown = [(g, h) for g, h in named if len(g[1]) > 1 or h]
    rest = [g for g, h in named if not (len(g[1]) > 1 or h)]
    w("**The formulas behind everything else.** Cells that say the same "
      "thing about their own position are one row here; the example is "
      "the first of them, in ordinary A1.")
    w("")
    w("| Cells | Named | Example | Formula |")
    w("|---|---|---|---|")
    for (_, cells), head in shown[:TABLE_CAP]:
        row, col, formula = cells[0]
        # A group may span several columns - two scaled copies dividing
        # by the same span cell say the same thing about their own
        # position - so it is named for all of them, not for whichever
        # came first.
        heads, seen = [], set()
        for c in sorted({c for _, c, _ in cells}):
            name = s.header_for(row, c)
            if name and name not in seen:
                seen.add(name)
                heads.append(name)
        w("| %s | %s | `%s` | `%s` |"
          % (ranges([(r, c) for r, c, _ in cells]),
             ", ".join(heads[:4]) or head or "-", addr(row, col),
             formula.replace("|", BAR)))
    if len(shown) > TABLE_CAP:
        w("")
        w("*%d further named shapes are not listed - this sheet has %d "
          "in all.*" % (len(shown) - TABLE_CAP, len(shown)))
    if rest:
        w("")
        w("%d unnamed one-off formula%s: %s."
          % (len(rest), "" if len(rest) == 1 else "s",
             ", ".join("`%s`" % addr(c[1][0][0], c[1][0][1])
                       for c in rest[:12])
             + (", ..." if len(rest) > 12 else "")))
    w("")

    if charts:
        w("**Chart objects**, in points from the top left of the sheet. The "
          "positions are what make separate charts read as one figure, so they "
          "are declared rather than dragged.")
        w("")
        w("| Name | Left | Top | Width | Height | Series | Axis |")
        w("|---|---|---|---|---|---|---|")
        for c in charts:
            axis = ("%g .. %g" % (c["min"], c["max"])) if "min" in c else "-"
            w("| `%s` | %g | %g | %g | %g | %s | %s |"
              % (c["name"], c["left"], c["top"], c["width"], c["height"],
                 c.get("series", "-"), axis))
        w("")
        if any("min" in c for c in charts):
            w("An axis range that is not round - `0.9` rather than `900` - is "
              "the scale block at work: the bounds were divided by the span "
              "cell, exactly as the values were. Multiply an axis bound by its "
              "span and you get back the range the geometry was measured in.")
            w("")
    return out


# ------------------------------------------------------------ checking it
#: A quoted formula in a technique section, with the address it names.
#: Built by concatenation rather than written out, because the pattern is
#: mostly newlines and an escaped newline inside a raw string is not one.
NL = chr(10)
QUOTED = re.compile("```" + NL + "(=[^" + NL + "]*)" + NL + "```"
                    + NL + NL + r"Read from `([^!`]+)!([A-Z]+\d+)`")
HEADING = re.compile(r"^### ([A-Za-z0-9 ]+)$", re.M)
ROW = re.compile(r"^\| [^|]*\| [^|]*\| `([A-Z]+\d+)` \| `(=[^`]*)` \|$", re.M)


def check_doc(md_path, xlsx_path) -> list[str]:
    """Assert every formula the document quotes is still in the workbook.

    Generating a document from the workbook stops it being *written* wrong. It
    does not stop it being *edited* wrong afterwards, or shipped beside a
    workbook that has since been rebuilt - and a document that quotes a formula
    the file does not contain is worse than no document, because it will be
    believed. This reads the markdown back and checks every quotation against
    the cell it names.
    """
    wb = openpyxl.load_workbook(str(xlsx_path))
    text = Path(md_path).read_text(encoding="utf-8")
    problems: list[str] = []
    checked = 0

    def compare(sheet, where, quoted):
        nonlocal checked
        if sheet not in wb.sheetnames:
            problems.append("%s: no such sheet, quoted at %s" % (sheet, where))
            return
        cell = wb[sheet][where]
        actual = clean(cell.value) if isinstance(cell.value, str) else ""
        checked += 1
        if actual != quoted.replace(BAR, "|"):
            problems.append("%s!%s: document says %r, workbook holds %r"
                            % (sheet, where, quoted, actual))

    for formula, sheet, where in QUOTED.findall(text):
        compare(sheet.strip(), where, formula)

    # The per-sheet tables inherit their sheet from the heading above them.
    marks = [(m.start(), m.group(1).strip()) for m in HEADING.finditer(text)]
    for i, (start, sheet) in enumerate(marks):
        end = marks[i + 1][0] if i + 1 < len(marks) else len(text)
        for where, formula in ROW.findall(text[start:end]):
            compare(sheet, where, formula)

    print("checked %d quoted formulas in %s" % (checked, os.path.basename(md_path)))
    return problems


def write_doc(xlsx_path, out_path, chart_facts=None, simple=False):
    text = build_markdown(str(xlsx_path), chart_facts, simple)
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(text, encoding="utf-8")
    return out_path


def main(argv=None):
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--xlsx", required=True, type=Path)
    ap.add_argument("--out", type=Path,
                    help="default: the workbook's name with a .md suffix")
    ap.add_argument("--simple", action="store_true")
    ap.add_argument("--no-charts", action="store_true",
                    help="skip the chart tables, which need Excel to read")
    ap.add_argument("--check", action="store_true",
                    help="do not write; verify that every formula the existing "
                         "document quotes is still in the workbook")
    a = ap.parse_args(argv)
    out = a.out or a.xlsx.with_suffix(".md")
    if a.check:
        problems = check_doc(out, a.xlsx)
        for problem in problems:
            print("  " + problem)
        print("document drift: %d" % len(problems) if problems
              else "document agrees with the workbook")
        return 1 if problems else 0
    facts = {} if a.no_charts else chart_facts_from(str(a.xlsx))
    print("wrote", write_doc(a.xlsx, out, facts, a.simple))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
