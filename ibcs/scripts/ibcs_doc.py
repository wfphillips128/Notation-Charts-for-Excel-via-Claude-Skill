"""Write the by-hand build guide for a workbook this skill just produced.

Not an inventory of formulas - a set of instructions. Somebody who has never
seen this project should be able to open a blank sheet, work through it, and
finish with the same chart, using nothing but Excel's own menus. The model is
``Vertical waterfall - manual steps.md``: what to type, where, which menu item,
what you should see, and what it looks like when it has gone wrong.

This module reads the facts. ``ibcs_guide.py`` writes the prose around them.

What keeps it honest
--------------------
Every concrete thing in the guide - a cell range, a formula, a hex colour, a
gap width, an axis bound, a chart's size in points - is **read back out of the
``.xlsx`` after it was written**, and out of the chart objects themselves while
Excel still has them open. The prose is written once; the numbers are never
typed twice. A hand-written guide drifts the first time a formula changes and
nothing catches it, because prose has no tests.

``--check`` runs the other way: it parses an existing guide, pulls out every
formula it quotes together with the address it names, and asserts the workbook
still holds exactly that. That is what stops a shipped document going stale
beside a rebuilt workbook.

Usage
-----
    python ibcs_excel.py --template C03A,C04A --out book.xlsx --doc book.md
    python ibcs_doc.py --xlsx book.xlsx --out book.md            # after the fact
    python ibcs_doc.py --xlsx book.xlsx --out book.md --check    # the gate
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

XL_VALUE, XL_CATEGORY = 2, 1

#: A pipe inside a markdown table cell has to be escaped.
BAR = chr(92) + chr(124)
NL = chr(10)

#: Rows of the per-sheet formula table before it stops listing. Whatever is
#: dropped is counted out loud - a table that silently truncates reads as
#: "that is all of them", which is the one thing it must not say.
TABLE_CAP = 40

#: Excel stores functions added after 2007 with an ``_xlfn.`` prefix, which is
#: a storage detail and not what anybody types.
XLFN = re.compile(r"_xlfn\.")

#: One A1 reference, with optional absolute markers.
REF = re.compile(r"(?<![A-Za-z0-9_])(\$?)([A-Z]{1,3})(\$?)([1-9][0-9]{0,6})"
                 r"(?![A-Za-z0-9_(])")

TECHNIQUE_PROBES = {
    "subject": lambda f: f.startswith("=") and '&" in "&' in f,
    "caption": lambda f: f.startswith('="') and "&$B$" in f,
    "variance": lambda f: bool(re.fullmatch(r"=[A-Z]+\d+-[A-Z]+\d+", f)),
    "relative": lambda f: "=IF(" in f and "/" in f and "*100)" in f,
    "split": lambda f: ">0," in f and "NA()" in f and f.count("IF") == 1,
    "scenario": lambda f: '="AC"' in f,
    "span": lambda f: "AGGREGATE(4,6" in f,
    "scaled": lambda f: bool(re.fullmatch(r"=[A-Z]+\d+/\$[A-Z]+\$\d+", f)),
    "text": lambda f: "TEXT(" in f and ("ISNA(" in f or "ISBLANK(" in f),
    "cascade": lambda f: bool(re.fullmatch(r"=[A-Z]+\d+\+[A-Z]+\d+\*[A-Z]+\d+",
                                           f)),
    "subtotal": lambda f: f.startswith("=SUM("),
    "clip": lambda f: "MEDIAN(" in f,
    "marker": lambda f: "UNICHAR(" in f,
}

#: The order technique examples are looked for in. Reading order would take
#: most of them off C01A, whose formulas carry an extra condition each and are
#: the worst place to meet an idea for the first time.
TEACHING_ORDER = ("C03A", "C04A", "C12A", "T04A")

#: MsoPatternType -> what the pattern is called in Excel's own Fill gallery.
PATTERNS = {26: "Wide upward diagonal", 25: "Wide downward diagonal",
            22: "Light upward diagonal", 21: "Light downward diagonal",
            28: "Dark upward diagonal", 27: "Dark downward diagonal"}

#: xlMarkerStyle -> the name in the Marker Options list.
MARKERS = {-4142: "None", 1: "Square", 2: "Diamond", 3: "Triangle",
           8: "Circle", 9: "Plus", -4168: "Star", 4: "X"}

#: xlChartType -> what you pick from the Insert Chart gallery.
CHART_TYPES = {51: "Clustered Column", 52: "Stacked Column",
               57: "Clustered Bar", 58: "Stacked Bar",
               4: "Line", 65: "Line with Markers",
               76: "Stacked Area", 75: "Area",
               -4169: "Scatter (X Y)", 15: "Bubble"}


# --------------------------------------------------------------- formatting
def clean(formula: str) -> str:
    """A formula as the user would type it."""
    return XLFN.sub("", formula or "")


def col_letters(col: int) -> str:
    s, c = "", int(col)
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def addr(row: int, col: int) -> str:
    return "%s%d" % (col_letters(col), row)


def hexrgb(value) -> str:
    """Excel stores colours as B<<16 | G<<8 | R, not the usual order.

    Getting this backwards is not a subtle bug - it turns the IBCS green into
    a blue - so the swap happens here once and nowhere else.
    """
    v = int(value)
    return "#%02X%02X%02X" % (v & 255, (v >> 8) & 255, (v >> 16) & 255)


def r1c1(formula: str, row: int, col: int) -> str:
    """The formula with every reference made relative to its own cell.

    Two cells hold "the same formula" when they say the same thing about their
    own position. It is the basis of the grouping in the per-sheet section -
    without it a twelve-row column reports as twelve findings.
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


def ranges(cells) -> str:
    """Summarise a set of (row, col) as an A1 range, honestly.

    A full rectangle is named as one. Anything else says how many cells it is
    and where they live, rather than claiming a shape it does not have.
    """
    cells = list(cells)
    rows = [r for r, _ in cells]
    cols = [c for _, c in cells]
    r0, r1, c0, c1 = min(rows), max(rows), min(cols), max(cols)
    box = (r1 - r0 + 1) * (c1 - c0 + 1)
    span = "%s:%s" % (addr(r0, c0), addr(r1, c1)) if box > 1 else addr(r0, c0)
    if len(cells) == box:
        return "`%s`" % span
    return "%d cells in `%s`" % (len(cells), span)


def plain(a1_range: str) -> str:
    return (a1_range or "").replace("$", "").split("!")[-1]


# ------------------------------------------- what the charts turned out to be
def _fill_of(fmt) -> str:
    """One shape's fill, named the way the Format pane names it."""
    try:
        fill = fmt.Fill
        if not fill.Visible:
            return "no fill"
        if int(fill.Type) == 2:                       # msoFillPatterned
            pattern = PATTERNS.get(int(fill.Pattern),
                                   "pattern %d" % int(fill.Pattern))
            return "%s, %s on %s" % (pattern, hexrgb(fill.ForeColor.RGB),
                                     hexrgb(fill.BackColor.RGB))
        return "solid %s" % hexrgb(fill.ForeColor.RGB)
    except Exception:                                         # noqa: BLE001
        return "?"


def _line_of(fmt) -> str:
    try:
        line = fmt.Line
        if not line.Visible:
            return "no line"
        return "%s at %.2gpt" % (hexrgb(line.ForeColor.RGB), float(line.Weight))
    except Exception:                                         # noqa: BLE001
        return "?"


def sheet_chart_facts(sheet) -> list:
    """One sheet's chart objects, as plain data. Safe to call mid-build.

    Read from the live chart rather than from the layout that asked for it, so
    the guide states what Excel *did* and not what it was told. Those differ
    more often than is comfortable - a plot-area assignment that did not take,
    a gap width Excel rounded, a marker style it declined.
    """
    facts = []
    for obj in sheet.ChartObjects():
        chart = obj.Chart
        d = {"name": obj.Name,
             "left": round(float(obj.Left), 1), "top": round(float(obj.Top), 1),
             "width": round(float(obj.Width), 1),
             "height": round(float(obj.Height), 1), "series": []}
        try:
            d["ctype"] = int(chart.ChartType)
        except Exception:                                     # noqa: BLE001
            pass
        for attr, key in (("InsideLeft", "pl"), ("InsideTop", "pt"),
                          ("InsideWidth", "pw"), ("InsideHeight", "ph")):
            try:
                d[key] = round(float(getattr(chart.PlotArea, attr)), 1)
            except Exception:                                 # noqa: BLE001
                pass
        try:
            group = chart.ChartGroups(1)
            d["gap"] = int(group.GapWidth)
            d["overlap"] = int(group.Overlap)
        except Exception:                                     # noqa: BLE001
            pass
        try:
            axis = chart.Axes(XL_VALUE)
            d["min"] = round(float(axis.MinimumScale), 6)
            d["max"] = round(float(axis.MaximumScale), 6)
            d["ticks"] = int(axis.TickLabelPosition) != -4142
        except Exception:                                     # noqa: BLE001
            pass
        try:
            n = int(chart.SeriesCollection().Count)
        except Exception:                                     # noqa: BLE001
            n = 0
        for i in range(1, n + 1):
            series = chart.SeriesCollection(i)
            s = {"n": i}
            for key, get in (
                    ("name", lambda: str(series.Name)),
                    ("ctype", lambda: int(series.ChartType)),
                    ("fill", lambda: _fill_of(series.Format)),
                    ("line", lambda: _line_of(series.Format)),
                    ("marker", lambda: MARKERS.get(int(series.MarkerStyle),
                                                   str(series.MarkerStyle))),
                    ("labels", lambda: bool(series.HasDataLabels))):
                try:
                    s[key] = get()
                except Exception:                             # noqa: BLE001
                    pass
            try:
                bars = series.ErrorBars
                s["stem"] = "%.2gpt" % float(bars.Format.Line.Weight)
                s["stem_colour"] = hexrgb(bars.Format.Line.ForeColor.RGB)
            except Exception:                                 # noqa: BLE001
                pass
            # How each point is painted. Two reasons this has to be per point
            # rather than per series, and both are notation:
            #
            #   * the forecast hatch is applied to the forecast months of an
            #     otherwise-actual series;
            #   * a marker series keeps its colours on the MARKER, and the pin
            #     heads carry a scenario each, so the series-level
            #     MarkerBackgroundColor is Excel's untouched default - reading
            #     it puts a theme blue in the guide where the sheet draws dark
            #     grey.
            marked = s.get("marker") not in (None, "None")
            try:
                overrides = {}
                for p in range(1, int(series.Points().Count) + 1):
                    point = series.Points(p)
                    if marked:
                        try:
                            key = "head %s, border %s" % (
                                hexrgb(point.MarkerBackgroundColor),
                                hexrgb(point.MarkerForegroundColor))
                        except Exception:                     # noqa: BLE001
                            continue
                    else:
                        key = _fill_of(point.Format)
                    overrides.setdefault(key, []).append(p)
                if overrides:
                    s["paint"] = sorted(overrides.items(),
                                        key=lambda kv: -len(kv[1]))
                if len(overrides) > 1 and not marked:
                    s["points"] = s["paint"]
            except Exception:                                 # noqa: BLE001
                pass
            d["series"].append(s)
        facts.append(d)
    return facts


def chart_facts_from(xlsx_path: str) -> dict:
    """Chart facts read out of a saved file, when the caller did not collect
    them during the build. Needs Excel; without it the guide omits the chart
    tables rather than guessing at them."""
    if win32 is None:
        return {}
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    out, wb = {}, None
    try:
        wb = excel.Workbooks.Open(os.path.abspath(xlsx_path), ReadOnly=True)
        for sheet in wb.Sheets:
            out[sheet.Name] = sheet_chart_facts(sheet)
    except Exception as exc:                                  # noqa: BLE001
        # Loudly. A guide that quietly loses every chart instruction because
        # a stray Excel instance was in the way still says "wrote" and looks
        # finished - which is how half a document goes missing unnoticed.
        print("warning: could not read the charts back (%s: %s). The guide "
              "will have no chart instructions in it; close any open Excel "
              "and run again." % (exc.__class__.__name__, exc),
              file=sys.stderr)
        return {}
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
            excel.Quit()
        except Exception:                                     # noqa: BLE001
            pass
    return out


# ------------------------------------------------------------ reading a sheet
class SheetFacts:
    """Everything one worksheet says about itself, read rather than assumed."""

    def __init__(self, ws):
        self.ws = ws
        self.name = ws.title
        self.formulas = {}
        self.typed = []
        self.labels = []
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
        out = []
        for shape, cells in self.formulas.items():
            cells.sort()
            out.append((shape, cells))
        out.sort(key=lambda g: (-len(g[1]), g[1][0]))
        return out

    def value(self, row, col):
        v = self.ws.cell(row=row, column=col).value
        return "" if v is None else (clean(v) if isinstance(v, str) else v)

    def header_for(self, row: int, col: int) -> str:
        """What a cell is called: its column header, or failing that its row.

        Only one row up is looked at - two rows up is a caption or a title.
        Length matters as well as position: the message line sits directly over
        the subject formula and is a paragraph, not a header, so anything long
        is rejected and the row label in column A is used instead.
        """
        above = self.ws.cell(row=max(1, row - 1), column=col).value
        if (isinstance(above, str) and above and not above.startswith("=")
                and len(above) <= 40):
            return above
        left = self.ws.cell(row=row, column=1).value
        if isinstance(left, str) and left and not left.startswith("="):
            return left[:40]
        return ""

    def find(self, test):
        best = None
        for cells in self.formulas.values():
            for row, col, f in cells:
                if test(f) and (best is None or (row, col) < best[:2]):
                    best = (row, col, f)
        return best

    def find_text(self, text: str):
        """(row, col) of the first cell whose text is exactly `text`."""
        for row, col in self.labels:
            if str(self.ws.cell(row=row, column=col).value).strip() == text:
                return row, col
        return None

    def header_row(self):
        """The data-zone header row, and the columns it spans.

        Found rather than declared: the widest run of adjacent text cells that
        has formulas or numbers under it. That is what a header row *is*, and
        it holds whichever family the sheet belongs to - which matters, because
        the seven families put their headers on different rows.
        """
        filled = {(r, c) for cells in self.formulas.values()
                  for r, c, _ in cells} | set(self.typed)
        by_row = {}
        for row, col in self.labels:
            by_row.setdefault(row, []).append(col)
        best = (0, None, [])
        for row, cols in sorted(by_row.items()):
            cols = sorted(cols)
            runs, run, start = [], 1, cols[0]
            for a, b in zip(cols, cols[1:]):
                if b == a + 1:
                    run += 1
                else:
                    runs.append((run, start))
                    run, start = 1, b
            runs.append((run, start))
            width, first = max(runs)
            under = sum(1 for c in range(first, first + width)
                        if (row + 1, c) in filled)
            if width > best[0] and under >= max(2, width // 3):
                best = (width, row, list(range(first, first + width)))
        return best[1], best[2]

    def data_rows(self, header):
        """First and last row of the block under a header row."""
        filled = {r for cells in self.formulas.values()
                  for r, _c, _f in cells} | {r for r, _c in self.typed}
        rows = sorted(r for r in filled if header is not None and r > header)
        if not rows:
            return None, None
        first, last = rows[0], rows[0]
        for r in rows:
            if r > last + 2:
                break
            last = r
        return first, last


def find_example(sheets, key):
    """One real example of a technique: (sheet name, address, formula)."""
    test = TECHNIQUE_PROBES[key]
    rank = {n: i for i, n in enumerate(TEACHING_ORDER)}
    for s in sorted(sheets, key=lambda x: rank.get(x.name, len(rank))):
        hit = s.find(test)
        if hit:
            return s.name, addr(hit[0], hit[1]), hit[2]
    return None


def quote(example, note: str = ""):
    """A found formula, rendered with the address it was found at.

    The shape of this block is what ``--check`` parses, so it is fixed.
    """
    if not example:
        return ["*(nothing in this workbook does this)*", ""]
    sheet, where, formula = example
    return ["```", formula, "```", "",
            "Read from `%s!%s`.%s" % (sheet, where,
                                      (" " + note) if note else ""), ""]


# ----------------------------------------------------------------- families
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


# ------------------------------------------------------------ checking it
#: A quoted formula in a walkthrough, with the address it names. Built by
#: concatenation rather than written out, because the pattern is mostly
#: newlines and an escaped newline inside a raw string is not one.
QUOTED = re.compile("```" + NL + "(=[^" + NL + "]*)" + NL + "```"
                    + NL + NL + r"Read from `([^!`]+)!([A-Z]+\d+)`")
HEADING = re.compile(r"^#{2,4} .*?([A-Z]\d{2}[A-Z])\b.*$", re.M)
ANY_HEADING = re.compile(r"^#{2,4} (.+)$", re.M)
ROW = re.compile(r"^\| [^|]*\| [^|]*\| `([A-Z]+\d+)` \| `(=[^`]*)` \|$", re.M)


def check_doc(md_path, xlsx_path):
    """Assert every formula the guide quotes is still in the workbook.

    Generating a guide from the workbook stops it being *written* wrong. It
    does not stop it being *edited* wrong afterwards, or shipped beside a
    workbook that has since been rebuilt - and a guide that quotes a formula
    the file does not contain is worse than no guide, because it will be
    believed and typed in.
    """
    wb = openpyxl.load_workbook(str(xlsx_path))
    text = Path(md_path).read_text(encoding="utf-8")
    problems, checked = [], [0]

    def compare(sheet, where, quoted):
        if sheet not in wb.sheetnames:
            problems.append("%s: no such sheet, quoted at %s" % (sheet, where))
            return
        cell = wb[sheet][where]
        actual = clean(cell.value) if isinstance(cell.value, str) else ""
        checked[0] += 1
        if actual != quoted.replace(BAR, "|"):
            problems.append("%s!%s: guide says %r, workbook holds %r"
                            % (sheet, where, quoted, actual))

    for formula, sheet, where in QUOTED.findall(text):
        compare(sheet.strip(), where, formula)

    # Per-sheet tables inherit their sheet from the nearest heading above them
    # that names one. A walkthrough heading names a sheet too, which is what
    # makes those tables checkable as well.
    marks = [(m.start(), m.group(1)) for m in HEADING.finditer(text)]
    for i, (start, sheet) in enumerate(marks):
        end = marks[i + 1][0] if i + 1 < len(marks) else len(text)
        for where, formula in ROW.findall(text[start:end]):
            compare(sheet, where, formula)

    print("checked %d quoted formulas in %s"
          % (checked[0], os.path.basename(str(md_path))))
    return problems


# ----------------------------------------------------------------- the doc
def build_markdown(xlsx_path: str, chart_facts=None, simple: bool = False):
    import ibcs_guide                                          # noqa: PLC0415
    wb = openpyxl.load_workbook(xlsx_path)
    sheets = [SheetFacts(wb[n]) for n in wb.sheetnames]
    return ibcs_guide.document(
        os.path.basename(str(xlsx_path)), sheets,
        chart_facts if chart_facts is not None else {}, simple)


def write_doc(xlsx_path, out_path, chart_facts=None, simple=False):
    text = build_markdown(str(xlsx_path), chart_facts, simple)
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(text, encoding="utf-8")
    return out_path


def main(argv=None):
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description=__doc__.split(NL)[0])
    ap.add_argument("--xlsx", required=True, type=Path)
    ap.add_argument("--out", type=Path,
                    help="default: the workbook's name with a .md suffix")
    ap.add_argument("--simple", action="store_true")
    ap.add_argument("--no-charts", action="store_true",
                    help="skip whatever needs Excel to read back")
    ap.add_argument("--check", action="store_true",
                    help="do not write; verify that every formula the existing "
                         "guide quotes is still in the workbook")
    a = ap.parse_args(argv)
    out = a.out or a.xlsx.with_suffix(".md")
    if a.check:
        problems = check_doc(out, a.xlsx)
        for problem in problems:
            print("  " + problem)
        print("document drift: %d" % len(problems) if problems
              else "the guide agrees with the workbook")
        return 1 if problems else 0
    facts = {} if a.no_charts else chart_facts_from(str(a.xlsx))
    print("wrote", write_doc(a.xlsx, out, facts, a.simple))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
