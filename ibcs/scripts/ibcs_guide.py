"""The prose of the by-hand build guide. ``ibcs_doc.py`` supplies the facts.

Everything concrete here - a range, a formula, a hex colour, a gap width, an
axis bound - is interpolated from what was read back out of the workbook. The
sentences are written once; the numbers are never typed twice.

The guide is in three parts, and the split is deliberate.

* **Part 1** is the skeleton every sheet in the library shares. Seven steps,
  and they are the same seven whichever chart you are building. Somebody who
  reads only this can already lay out a sheet that a chart can be hung on.
* **Part 2** is a worked build. The tier stack is done at full depth - every
  menu, every setting, what you should see after each step - because it carries
  the notation and because the other families are variations on it. The rest
  are written as *what differs*, which is the honest length for them and keeps
  the shared steps in one place instead of six.
* **Part 3** is per sheet: its ranges, its charts, and its formulas grouped by
  shape, for when you are rebuilding one particular template rather than
  learning the method.
"""

from __future__ import annotations

import ibcs_doc as F

NL = chr(10)

FAMILY_ORDER = ("tier stack", "table", "structure", "line", "XY", "tree",
                "panel")

FAMILY_HEAD = {
    "tier stack": "a tier stack",
    "table": "a table with drawn columns",
    "structure": "a structure chart",
    "line": "a line sheet",
    "XY": "an XY plot",
    "tree": "a driver tree",
    "panel": "a panel grid",
}


def _fmt(x):
    """A number as a human would write it."""
    if isinstance(x, float):
        return ("%.4f" % x).rstrip("0").rstrip(".") if abs(x) < 1000 \
            else "%,.0f".replace(",", "") % x
    return str(x)


def _chart(charts, name):
    for c in charts:
        if c.get("name") == name:
            return c
    return None


def _ctype(c):
    return F.CHART_TYPES.get(c.get("ctype"), "chart type %s" % c.get("ctype"))


def _paint(series):
    """How a series' points are painted, as one cell of a table.

    Per point rather than per series, because both of the things that vary -
    the forecast hatch and the pin head's scenario - vary within a series.
    """
    groups = series.get("paint") or []
    if not groups:
        return series.get("fill", "-")
    out = []
    for how, points in groups[:3]:
        span = ("%d-%d" % (points[0], points[-1])
                if points == list(range(points[0], points[-1] + 1))
                else ", ".join(str(p) for p in points[:6]))
        # "points 10-12: hatched" rather than "hatched on 10-12", which read
        # as a second "on" after the fill's own "#404040 on #F2F2F2".
        out.append("points %s: %s" % (span, how) if len(groups) > 1 else how)
    return "; ".join(out)


def _kind(sheet, row, col):
    """Whether a column is typed or derived, said the way a reader needs it.

    A category name and a scenario code are typed as surely as a number is -
    calling them "formula" because they are not numeric would send somebody
    hunting for a formula that is not there.
    """
    if row is None:
        return "?"
    value = sheet.value(row, col)
    if isinstance(value, str) and value.startswith("="):
        return "formula"
    if isinstance(value, str) and value:
        return "**typed** (text)"
    if value == "" or value is None:
        return "-"
    return "**typed**"


# --------------------------------------------------------------------------- #
# Part 0 - the front matter
# --------------------------------------------------------------------------- #
def _front(w, book, templates, charts, simple):
    w("# Building `%s` by hand in Excel" % book)
    w("")
    w("This is the same workbook, done manually - no macros, no add-ins, no "
      "code. Work through Part 1 once and you have the sheet every template in "
      "the library sits on; Part 2 builds the charts on top of it.")
    w("")
    w("**Every range, formula, colour, gap width and axis bound below was read "
      "back out of `%s` after it was written.** Where this says a formula sits "
      "in `C03A!E13`, it does; where it says a fill is `#404040`, that is the "
      "value in the file. The prose was written once. The numbers are not "
      "typed twice, which is the only way a document like this stays true to "
      "the thing it describes." % book)
    w("")
    if simple:
        w("This is the guide to the **simple** workbook - each template reduced "
          "to its base tiers. The construction is identical to the full "
          "workbook's; there is simply less of it on each sheet.")
        w("")

    w("## What you need")
    w("")
    w("| | |")
    w("|---|---|")
    w("| Excel | 2016 or later, Windows or Mac. Nothing here needs 365 |")
    w("| Time | about an hour for your first tier stack, twenty minutes after |")
    w("| Skills | writing a formula, inserting a chart, and `Ctrl+1` |")
    w("")
    w("`Ctrl+1` opens the Format pane for whatever is selected, and it is the "
      "single most useful key in this whole exercise. `Cmd+1` on a Mac.")
    w("")
    w("Two habits will save you most of the pain:")
    w("")
    w("- **Use the Current Selection dropdown.** Chart Format tab, far left. "
      "Series in these charts are deliberately thin, invisible, or stacked "
      "under each other, and hunting for them with the mouse is miserable. "
      "Pick them from the list by name instead.")
    w("- **Name every chart as you create it.** Click the chart, then type the "
      "name into the Name Box (left of the formula bar) and press Enter. The "
      "names used below are the ones in the workbook, and by the time you have "
      "three overlapping charts on a sheet you will want them.")
    w("")

    w("## The one idea")
    w("")
    w("**A chart in this notation is several charts, stacked, sharing one "
      "category axis.**")
    w("")
    w("A measure and a variance are different units and different scales, so "
      "they cannot share a value axis - but they describe the same twelve "
      "months, so they must line up. Excel will not do that inside one chart "
      "object. So each tier is its own chart, sized and positioned in points "
      "so that category 1 of the top tier sits directly above category 1 of "
      "the bottom one.")
    w("")
    w("```")
    w("   Delta PY%    |  pins, own scale, own chart object")
    w("   Delta PY     |  columns, own scale, own chart object")
    w("   Contribution |  columns, own scale, own chart object")
    w("                 Jan Feb Mar Apr ...   <- one shared category axis")
    w("```")
    w("")
    w("Everything else in this guide is in service of that, or of the second "
      "idea: **the reader edits cells, never the chart.**")
    w("")

    w("## The notation you are reproducing")
    w("")
    w("IBCS is a notation, not a style guide: the same meaning must always "
      "take the same visual form. Four fills carry most of it, and these are "
      "the values in this workbook:")
    w("")
    w("| Scenario | Means | Fill | Border |")
    w("|---|---|---|---|")
    w("| AC | actual - it happened | solid `#404040` | none |")
    w("| PY | prior year - it happened, earlier | solid `#A6A6A6` | none |")
    w("| PL / BU | plan or budget - fictitious | **no fill** | `#404040` |")
    w("| FC | forecast - expected | **Wide upward diagonal**, `#404040` on "
      "`#F2F2F2` | `#404040` |")
    w("")
    w("| Variance | Means | Fill |")
    w("|---|---|---|")
    w("| good | favourable impact | `#8CB400` |")
    w("| bad | adverse impact | `#FF0000` |")
    w("| neutral | no direction to it | `#0064FF` |")
    w("")
    w("> **Colour is by impact, not by sign.** A cost 40 over plan is a "
      "positive number and a bad outcome, so it is red. If you colour on "
      "`>0` you will get every cost overrun green, and the chart will be "
      "confidently wrong.")
    w("")
    w("> **IBCS prescribes no colour codes.** Rule UN 4.1 says so outright. "
      "These hex values are this project's house palette, not a requirement of "
      "the standard.")
    w("")

    w("## What is in this workbook")
    w("")
    w("| Sheet | What it is | Charts |")
    w("|---|---|---|")
    for s in templates:
        subject = s.value(2, 2) or ""
        w("| [%s](#%s) | %s%s | %d |"
          % (s.name, s.name.lower(), FAMILY_HEAD[F.family_of(s.name)],
             (" - " + str(subject).lower()) if subject else "",
             len(charts.get(s.name, []))))
    w("")
    w("---")
    w("")


# --------------------------------------------------------------------------- #
# Part 1 - the skeleton
# --------------------------------------------------------------------------- #
def _skeleton(w, templates, charts):
    ref = next((s for s in templates if s.name in F.TEACHING_ORDER), templates[0])
    hdr_row, hdr_cols = ref.header_row()
    first, last = ref.data_rows(hdr_row)

    w("# Part 1 - The skeleton every sheet shares")
    w("")
    w("Seven steps. They are the same seven whichever chart you are going to "
      "build on top, so they are here once rather than seven times. `%s` is "
      "used for the addresses; every other sheet is the same shape with "
      "different columns." % ref.name)
    w("")

    # --- 1 -----------------------------------------------------------------
    w("## Step 1 - Six typed lines at the top, and nothing else typed")
    w("")
    w("Rows 1 to 6 of column A are labels, and column B is the only text "
      "anybody ever types:")
    w("")
    w("| Row | A | B |")
    w("|---|---|---|")
    for r in range(1, 7):
        text = str(ref.value(r, 2))
        if len(text) > 96:
            text = text[:93].rsplit(" ", 1)[0] + " ..."
        w("| %d | %s | %s |" % (r, ref.value(r, 1), text))
    w("")
    w("Row 7 is the **subject line**, and it is a formula:")
    w("")
    w.extend(F.quote(F.find_example(templates, "subject")))
    w("Then the tier captions, one per drawn tier that needs one:")
    w("")
    w.extend(F.quote(F.find_example(templates, "caption")))
    w("This is the part people skip, and it is the part that makes the sheet "
      "worth keeping. **Nothing a chart says may be typed into the chart.** "
      "Every text box on the finished chart is *linked* to one of these cells, "
      "so retyping `kEUR` as `kGBP` in `B3` changes the subject line and every "
      "caption at once. Type it into the chart and it is a lie the first time "
      "the data changes.")
    w("")
    w("To link a text box: draw it, then with the box selected click in the "
      "**formula bar**, type `=` and click the cell. Not into the box - into "
      "the formula bar. It is the one place in Excel where that is the whole "
      "technique.")
    w("")

    # --- 2 -----------------------------------------------------------------
    w("## Step 2 - Two zones, and nothing straddles them")
    w("")
    w("The sheet is a **data zone** on the left of declared width, and the "
      "**charts** to its right. The print area is the chart zone alone - the "
      "data is the input, not the deliverable.")
    w("")
    w("| Sheet | Cells used | Print area |")
    w("|---|---|---|")
    for s in templates:
        w("| %s | `%s` | `%s` |" % (s.name, s.ws.dimensions,
                                    F.plain(s.ws.print_area) or "-"))
    w("")
    tables = [s.name for s in templates if F.family_of(s.name) == "table"]
    if tables:
        w("A **table** sheet is the other way round: on %s the grid *is* the "
          "deliverable, so the print area is the grid, starting below the six "
          "typed title rows. Same rule - the print area holds what you are "
          "publishing and the typed inputs are not it."
          % ", ".join(tables))
        w("")
    w("> **Never AutoFit a column a chart is positioned against.** The charts "
      "are placed in points, measured from where the data zone ends. One long "
      "customer name and AutoFit moves that boundary, and the charts end up "
      "drawn on top of the source data. Set column widths by hand: "
      "**Home -> Format -> Column Width**.")
    w("")

    # --- 3 -----------------------------------------------------------------
    w("## Step 3 - The data block")
    w("")
    if hdr_row:
        w("On `%s` the headers are on row **%d** (`%s`) and the data runs "
          "rows **%d to %d**."
          % (ref.name, hdr_row, F.addr(hdr_row, hdr_cols[0]) + ":"
             + F.addr(hdr_row, hdr_cols[-1]), first, last))
        w("")
        w("| Column | Header | Typed or derived |")
        w("|---|---|---|")
        for c in hdr_cols[:14]:
            head = ref.value(hdr_row, c)
            w("| %s | %s | %s |" % (F.col_letters(c), head,
                                    _kind(ref, first, c)))
        if len(hdr_cols) > 14:
            w("| ... | %d more, all formulas | |" % (len(hdr_cols) - 14))
        w("")
    w("Two columns earn special mention because they are notation rather than "
      "data:")
    w("")
    w("- **A scenario column.** One of `AC`, `PY`, `PL`, `BU`, `FC` per row. "
      "This is what later decides which fill a bar gets, so it is a *value* in "
      "a cell and not a colour somebody applied. That is the whole reason a "
      "forecast can turn into an actual by typing over one cell.")
    w("- **A sign or kind column**, on the waterfall templates: `+1` / `-1`, "
      "or `total` / `state` / `variance`. The cascade formulas read it.")
    w("")
    w("Shade the typed cells so a reader can see what is theirs to edit, and "
      "leave everything else unshaded - including the subtotals, which is the "
      "point.")
    w("")

    # --- 4 -----------------------------------------------------------------
    w("## Step 4 - Split every drawn series by what it means")
    w("")
    w("This is the step that surprises people, and everything downstream "
      "depends on it.")
    w("")
    w("**Excel gives one colour to a series.** So a variance that is green "
      "when favourable and red when adverse cannot be one series - it has to "
      "be two, split in the *cells*, each holding `NA()` where the other owns "
      "the point:")
    w("")
    w.extend(F.quote(F.find_example(templates, "split")))
    w("Do the same wherever a fill changes for any reason: actual against "
      "forecast, plan against budget, an increase against a decrease.")
    w("")
    w("> ### `NA()`, never `\"\"`")
    w(">")
    w("> An empty string is **not** an empty cell. Excel plots it as a **zero**, "
      "so every gap becomes a bar of nothing sitting on the axis. `NA()` is a "
      "genuine gap. Also set **Select Data -> Hidden and Empty Cells -> Show "
      "empty cells as: Gaps**.")
    w(">")
    w("> This is the single most common way to end up with a chart that looks "
      "broken for no visible reason.")
    w("")
    w("And once more, because it is the difference between a correct chart and "
      "a confident lie: **split by impact, not by sign.** Within one variance "
      "column of a P&L the favourable values sit on both sides of the axis, "
      "because a cost line going up is adverse and a cost line going down is "
      "favourable.")
    w("")

    # --- 5 -----------------------------------------------------------------
    w("## Step 5 - The scale block, so the charts follow *your* numbers")
    w("")
    w("Skip this one and everything still works - until somebody pastes their "
      "own figures in, and every bar clips or shrinks to nothing.")
    w("")
    w("**Excel will not bind an axis bound to a formula.** An axis maximum is "
      "a number you type into a box; it cannot be `=MAX(...)`. So a chart "
      "built for one set of figures is built for *those* figures.")
    w("")
    w("The way round it is to turn the problem over. Instead of scaling the "
      "axis to the data, **scale the data to a fixed axis**. Three parts:")
    w("")
    w("**1. A span cell** - one per group of tiers that share a unit - holding "
      "the largest magnitude anywhere in that group:")
    w("")
    w.extend(F.quote(F.find_example(templates, "span"),
                     "`AGGREGATE(4,6,...)` is MAX ignoring errors, which "
                     "matters because half of these columns are deliberately "
                     "`NA()`."))
    w("**2. A scaled copy of every column a chart reads**, which is what you "
      "actually plot:")
    w("")
    w.extend(F.quote(F.find_example(templates, "scaled")))
    w("**3. Axis bounds divided by the same span.** If your measure tier was "
      "designed to run 0 to 230, you type `0` and `230/span` - or rather, you "
      "work out that number once and type the result.")
    w("")
    w("Dividing both the values and the bounds by one number is a visual "
      "no-op at today's figures and follows the data at any others. And "
      "because zero divided by anything is zero, the zero line, the reference "
      "rules and every caption positioned from plot geometry stay exactly "
      "where they were.")
    w("")
    w("Hide the scaled columns - **right-click the column headers -> Hide**. "
      "A hidden column measures zero width, so you can add all of this to a "
      "sheet without moving a single chart on it.")
    w("")
    w("> **Where this cannot be used.** Every chart in this library hides its "
      "value axis, which is why dividing by a span is invisible. The two XY "
      "sheets are the exception - they show gridlines and tick labels, and "
      "dividing those by a span would print `0.25` where the data says "
      "`29.16%`. Those two get their axis fitted to the data once, at build "
      "time, and do not follow an edit. If you are building a chart whose axis "
      "a reader actually reads, do the same.")
    w("")

    # --- 6 -----------------------------------------------------------------
    w("## Step 6 - A text column for every label")
    w("")
    w("A data label linked to a cell shows the **cell's** value and ignores "
      "the label's own number format. Link a label to a raw variance and you "
      "get `31.61057692` printed beside a chart that says `+32`.")
    w("")
    w("So the formatting happens in a cell:")
    w("")
    w.extend(F.quote(F.find_example(templates, "text")))
    w("One such column per labelled series, blank where that series has no "
      "point. Then in step 2.9 you point the labels at these columns with "
      "**Value From Cells**.")
    w("")
    w("This is also why you cannot skip it: the label range applies to the "
      "whole series, so a shared column would print the other series' numbers "
      "on your bars.")
    w("")

    # --- 7 -----------------------------------------------------------------
    w("## Step 7 - Page setup, before you forget")
    w("")
    w("1. Select the range the charts cover, then **Page Layout -> Print Area "
      "-> Set Print Area**.")
    w("2. **Page Layout -> Orientation**, then **Scale to Fit: Width 1 page, "
      "Height 1 page**.")
    w("3. **View -> uncheck Gridlines**, so the worksheet grid does not print "
      "behind the charts.")
    w("4. Check nothing else is inside the print area. Cell contents under a "
      "chart print straight through it.")
    w("")
    w("---")
    w("")


# --------------------------------------------------------------------------- #
# Part 2 - the worked tier-stack build
# --------------------------------------------------------------------------- #
def _tier_walkthrough(w, sheet, charts, templates):
    name = sheet.name
    hdr_row, hdr_cols = sheet.header_row()
    first, last = sheet.data_rows(hdr_row)
    n = (last - first + 1) if first else 0

    w("## Building a tier stack, on `%s`" % name)
    w("")
    subject = sheet.value(2, 2)
    w("The full walkthrough. Every other family in Part 2 is written as *what "
      "differs from this*, so read it even if the chart you want is a "
      "different shape.")
    w("")
    if subject:
        w("`%s` is %s over %d categories, in %d tier%s."
          % (name, str(subject).lower(), n, len(charts),
             "" if len(charts) == 1 else "s"))
        w("")

    # --- the columns -------------------------------------------------------
    w("### 2.1 The columns, and what each is for")
    w("")
    w("Headers on row **%d**, data rows **%d-%d**. Typed cells are %s - "
      "everything else is a formula."
      % (hdr_row or 0, first or 0, last or 0,
         F.ranges(sheet.typed) if sheet.typed else "none"))
    w("")
    w("| Column | Header | Formula in row %d |" % (first or 0))
    w("|---|---|---|")
    shown = 0
    for c in range(1, sheet.ws.max_column + 1):
        head = sheet.value(hdr_row, c) if hdr_row else ""
        cell = sheet.value(first, c) if first else ""
        if not head and not cell:
            continue
        if str(cell).startswith("="):
            body = "`%s`" % str(cell).replace("|", F.BAR)
        else:
            body = "**typed** - `%s`, and down" % cell
        w("| %s | %s | %s |" % (F.col_letters(c), head, body))
        shown += 1
        if shown >= 34:
            w("| ... | | |")
            break
    w("")
    w("Enter each formula in row %d and fill down to row %d." % (first, last))
    w("")

    # --- what you should see ----------------------------------------------
    w("### 2.2 Check the numbers before you go near a chart")
    w("")
    w("If these are right the chart cannot go far wrong; if they are wrong no "
      "amount of formatting will save it.")
    w("")
    cols = [c for c in range(1, min(sheet.ws.max_column, 8) + 1)
            if sheet.value(hdr_row, c) or sheet.value(first, c)]
    w("| Row | " + " | ".join(str(sheet.value(hdr_row, c)) for c in cols) + " |")
    w("|---" * (len(cols) + 1) + "|")
    for r in range(first, min(first + 4, last + 1)):
        cells = []
        for c in cols:
            v = sheet.value(r, c)
            cells.append("`%s`" % str(v).replace("|", F.BAR)
                         if str(v).startswith("=") else str(v))
        w("| %d | %s |" % (r, " | ".join(cells)))
    w("| ... | " + " | ".join("" for _ in cols) + " |")
    w("")

    # --- insert the measure tier ------------------------------------------
    base = charts[-1] if charts else None
    if base:
        w("### 2.3 Insert the bottom tier")
        w("")
        w("The measure tier is the one everything else is aligned to, so it "
          "goes first.")
        w("")
        w("1. Select the category labels and the scaled measure columns "
          "together - hold **Ctrl** for the second block, and include the "
          "header row so the series get their names.")
        w("2. **Insert -> %s.**" % _ctype(base))
        w("3. Name it `%s` in the Name Box." % base["name"])
        w("4. `Ctrl+1` on the chart area -> **Size**: width **%g pt**, height "
          "**%g pt**. Under **Properties**, set **Don't move or size with "
          "cells** - otherwise inserting a row later moves your chart."
          % (base.get("width", 0), base.get("height", 0)))
        w("")
        w("Then the settings that are not defaults:")
        w("")
        w("| Where | Setting | Value |")
        w("|---|---|---|")
        if "gap" in base:
            w("| any series -> Series Options | Gap Width | **%d%%** |"
              % base["gap"])
        if "overlap" in base:
            w("| same pane | Series Overlap | **%d%%** |" % base["overlap"])
        if "min" in base:
            w("| value axis -> Axis Options | Minimum / Maximum | **%s** / "
              "**%s** |" % (_fmt(base["min"]), _fmt(base["max"])))
        w("| value axis -> Labels | Label Position | **None** |")
        w("| value axis -> Line | | **No line**, no tick marks |")
        w("| chart area -> Border | | **No line** |")
        w("| Chart Elements (+) | Gridlines, Legend, Title | **all off** |")
        w("")
        w("> The axis bounds are not round numbers because of the scale block "
          "in Part 1 step 5: they are the range the geometry was designed in, "
          "divided by the span cell. Multiply them back by the span and you "
          "get the round numbers you started from.")
        w("")

        # --- fills ---------------------------------------------------------
        w("### 2.4 The fills - this is the notation")
        w("")
        w("Select each series from **Chart Format -> Current Selection**, then "
          "`Ctrl+1` -> Fill & Line:")
        w("")
        w("| Series | Fill | Line |")
        w("|---|---|---|")
        for s in base["series"]:
            w("| `%s` | %s | %s |" % (s.get("name", "?"), s.get("fill", "?"),
                                      s.get("line", "?")))
        w("")
        hatched = [(s, o) for s in base["series"] for o in [s.get("points")]
                   if o]
        if hatched:
            s, overrides = hatched[0]
            w("**The forecast is per point, not per series.** On `%s` the "
              "series carries the actual fill and the forecast months are "
              "overridden individually:" % s.get("name", "?"))
            w("")
            w("| Points | Fill |")
            w("|---|---|")
            for fill, pts in overrides:
                span = ("%d-%d" % (pts[0], pts[-1])
                        if pts == list(range(pts[0], pts[-1] + 1))
                        else ", ".join(str(p) for p in pts[:8]))
                w("| %s | %s |" % (span, fill))
            w("")
            w("To do it: click the series once to select all of it, then click "
              "**again** on the single bar you want - that selects the point - "
              "then `Ctrl+1` -> Fill -> **Pattern fill**, and pick the pattern "
              "from the gallery with the foreground and background above.")
            w("")
            w("> Yes, this is manual, and yes it is the one thing on the sheet "
              "that does not follow the data. If the forecast starts a month "
              "earlier next quarter you must re-apply it. The alternative is a "
              "separate series for the forecast points, driven by the scenario "
              "column - more columns, no clicking. Both are used in this "
              "workbook; look at which sheets have an `AC` and an `FC` series "
              "and which have one series with overridden points.")
            w("")

    # --- the variance tier -------------------------------------------------
    if len(charts) > 1:
        var = charts[len(charts) - 2]
        w("### 2.5 The variance tier above it")
        w("")
        w("Same construction, different scale, and it is a **separate chart "
          "object**. Insert it exactly as in 2.3, from the two variance "
          "columns you split in Part 1 step 4.")
        w("")
        w("| | |")
        w("|---|---|")
        w("| Name | `%s` |" % var["name"])
        w("| Type | %s |" % _ctype(var))
        w("| Size | %g x %g pt |" % (var.get("width", 0), var.get("height", 0)))
        if "min" in var:
            w("| Axis | %s to %s |" % (_fmt(var["min"]), _fmt(var["max"])))
        if "gap" in var:
            w("| Gap width | %d%% |" % var["gap"])
        if "overlap" in var:
            w("| Overlap | %d%% |" % var["overlap"])
        w("")
        w("| Series | Fill |")
        w("|---|---|")
        for s in var["series"]:
            w("| `%s` | %s |" % (s.get("name", "?"), s.get("fill", "?")))
        w("")
        w("The invisible one is not a mistake. A variance drawn from a "
          "baseline that is not zero needs a transparent segment underneath "
          "holding the offset - the same trick a waterfall uses. Where the "
          "variance runs from zero, that series is still there and still zero, "
          "so the two tiers stay interchangeable.")
        w("")

    # --- pins --------------------------------------------------------------
    pin = next((c for c in charts
                if any("stem" in s for s in c.get("series", []))), None)
    if pin:
        w("### 2.6 The relative-variance tier: pins")
        w("")
        w("A relative variance is drawn as a **pin** - a thin stem with a head "
          "marker - and there is no pin chart type. The obvious construction, "
          "a very narrow column, does not work: `Gap Width` caps at 500%, "
          "which on a twelve-category axis bottoms out around a 9px stem where "
          "the reference draws 5.")
        w("")
        w("What works is a **line chart with markers, no line, and a custom Y "
          "error bar as the stem**. An error bar's weight is in *points*, so "
          "the stem is exactly as thin as you ask for.")
        w("")
        w("1. **Insert -> %s** from the two split relative-variance columns."
          % _ctype(pin))
        w("2. Each series -> `Ctrl+1` -> Fill & Line -> **Line: No line**. The "
          "markers are the pins; the line between them means nothing.")
        head = pin["series"][0]
        w("3. Marker Options -> **%s**, size to taste, and set **Marker Fill** "
          "and **Marker Border** - not the series fill, which on a "
          "marker-only series is not what you are looking at. The head "
          "carries the *minuend's* scenario, which is what says whether a "
          "relative variance was measured or forecast."
          % head.get("marker", "Square"))
        w("4. With the series selected, **Chart Elements (+) -> Error Bars -> "
          "More Options**. Direction **Minus** for the series that goes up and "
          "**Plus** for the one that goes down, End Style **No Cap**, then "
          "**Custom -> Specify Value** and point *both* boxes at that series' "
          "stem-length column.")
        w("5. The error bar -> `Ctrl+1` -> Line -> width **%s**."
          % next((s.get("stem") for s in pin["series"] if s.get("stem")), "?"))
        w("")
        w("| Series | Marker | Stem | Stem colour | Heads |")
        w("|---|---|---|---|---|")
        for s in pin["series"]:
            w("| `%s` | %s | %s | %s | %s |"
              % (s.get("name", "?"), s.get("marker", "?"),
                 s.get("stem", "-"), s.get("stem_colour", "-"),
                 _paint(s)))
        w("")
        w("The two stems are different colours because the direction of a "
          "variance is its meaning, and an error bar takes one colour for a "
          "whole series - which is *why* there are two series here rather "
          "than one.")
        w("")
        w("Four things about error bars are not in the documentation and each "
          "one fails quietly:")
        w("")
        w("- The stem-length column must be **entirely numeric** - write `0` "
          "where the series has no point, never `NA()`. An error-bar range "
          "with one non-numeric cell is discarded whole.")
        w("- **End Style must be No Cap.** The cap is the little crossbar; "
          "with it on, your pin has a T on the end.")
        w("- The **direction** is the counter-intuitive one. An up-pin's stem "
          "runs *down* from the marker to zero, so it is a **Minus** error bar.")
        w("- On a line series, clicking a point and setting its **Line** "
          "colour styles the *connecting segment*, not the marker border. Use "
          "**Marker Border** for the border.")
        w("")

    # --- alignment ---------------------------------------------------------
    w("### 2.7 Line the tiers up - the step that makes it one figure")
    w("")
    w("Three charts that nearly line up look like a mistake. They have to be "
      "exact, and you get there by typing numbers, not by dragging.")
    w("")
    w("For each chart: `Ctrl+1` -> **Size & Properties**, and set Height, "
      "Width, and under Position the Horizontal and Vertical offsets:")
    w("")
    w("| Chart | Left | Top | Width | Height |")
    w("|---|---|---|---|---|")
    for c in charts:
        w("| `%s` | %g | %g | %g | %g |"
          % (c["name"], c.get("left", 0), c.get("top", 0),
             c.get("width", 0), c.get("height", 0)))
    w("")
    w("That is not enough on its own. Two charts of the same width can still "
      "have plot areas of different widths, because Excel sizes the plot area "
      "around whatever labels it has to fit. So set the **plot area** too: "
      "click inside the plot (not the chart), `Ctrl+1`, and set its size and "
      "position:")
    w("")
    w("| Chart | Plot left | Plot top | Plot width | Plot height |")
    w("|---|---|---|---|---|")
    for c in charts:
        if "pl" in c:
            w("| `%s` | %g | %g | %g | %g |"
              % (c["name"], c.get("pl", 0), c.get("pt", 0),
                 c.get("pw", 0), c.get("ph", 0)))
    w("")
    w("> **Set the plot area size before the position, and check it after.** "
      "Excel treats a plot-area assignment as a request: it will quietly "
      "adjust what you asked for to fit the labels, and it applies width and "
      "position in an order that means setting one can undo the other. Set "
      "them, then look at the numbers again.")
    w("")
    w("Only the bottom tier keeps its category axis labels. On the others, "
      "select the category axis -> Labels -> **Label Position: None**. The "
      "labels are shared, so they are printed once.")
    w("")

    # --- captions ----------------------------------------------------------
    w("### 2.8 Captions and the title block")
    w("")
    w("Draw a text box for each of the four title lines and each tier caption, "
      "and link every one of them to its cell as in Part 1 step 1 - select the "
      "box, click in the formula bar, type `=`, click the cell.")
    w("")
    w("Set **Word Wrap off** on the captions, and turn off **Resize shape to "
      "fit text**. A caption that rewraps when the unit changes from `kEUR` to "
      "`kUSD` will push itself over the chart.")
    w("")

    # --- labels ------------------------------------------------------------
    w("### 2.9 The data labels")
    w("")
    w("For each series that carries labels:")
    w("")
    w("1. Select the series -> **Chart Elements (+) -> Data Labels**.")
    w("2. `Ctrl+1` -> Label Options -> tick **Value From Cells**, and select "
      "that series' text column from Part 1 step 6.")
    w("3. **Untick Value**, and untick everything else. Only Value From Cells "
      "stays on.")
    w("4. Position: **Outside End** for columns that grow from an axis, "
      "**Inside End** where the label would otherwise leave the plot.")
    w("")
    w("| Series | Labelled |")
    w("|---|---|")
    for c in charts:
        for s in c.get("series", []):
            if s.get("labels"):
                w("| `%s` / `%s` | yes |" % (c["name"], s.get("name", "?")))
    w("")

    # --- checks ------------------------------------------------------------
    w("### 2.10 Check it before you send it")
    w("")
    w("Numbers first, appearance second.")
    w("")
    w("- [ ] **Every scenario has the right fill.** Solid dark for actual, "
      "solid light for prior year, outlined for plan, hatched for forecast. "
      "One wrong fill and the chart says something untrue.")
    w("- [ ] **Variance colour follows impact, not sign.** Check a cost line "
      "that went up: it must be red.")
    w("- [ ] **The tiers line up.** Put a ruler on the screen - category 1 of "
      "the top tier over category 1 of the bottom.")
    w("- [ ] **The variance axis carries its reference.** A reader must be "
      "able to tell what the variance is *from* without a legend.")
    w("- [ ] **No stray zero labels** where a series has no point - that is "
      "the `\"\"`-instead-of-`NA()` symptom.")
    w("- [ ] **Every label is linked**, not typed. Change a number and watch "
      "them all move.")
    w("- [ ] **Retype the unit in `B3`.** The subject line and every caption "
      "must follow. If one does not, it was typed into the chart.")
    w("- [ ] **Nothing prints under the charts.** Print Preview, one page.")
    w("")

    # --- gotchas -----------------------------------------------------------
    w("### 2.11 What goes wrong, and why")
    w("")
    w("| Symptom | Cause | Fix |")
    w("|---|---|---|")
    w("| A bar of nothing sitting on the axis | a formula returns `\"\"` where "
      "it means `NA()` | step 4 |")
    w("| Every cost overrun is green | the split is on sign, not impact | "
      "step 4 |")
    w("| Bars clip the moment new data is pasted | axis bounds typed as data "
      "values with no scale block | step 5 |")
    w("| A label prints `31.61057692` | linked to the raw cell, not a `TEXT()` "
      "column | step 6 |")
    w("| Labels stop updating | typed into the chart instead of linked | "
      "step 1 |")
    w("| Tiers nearly line up | plot area sizes not set, only chart sizes | "
      "step 2.7 |")
    w("| Charts drawn over the data | AutoFit moved the zone boundary | "
      "step 2 |")
    w("| A pin has a T on the end | error-bar End Style is not No Cap | "
      "step 2.6 |")
    w("| Stray lines between pin heads | a point's Line was styled instead of "
      "its Marker Border | step 2.6 |")
    w("| An error bar does nothing | its range contains an `NA()` or a blank | "
      "step 2.6 |")
    w("| The whole chart re-scales when a row is inserted | chart Properties "
      "left on Move and size with cells | step 2.3 |")
    w("")


# --------------------------------------------------------------------------- #
# Part 2 - the shorter family walkthroughs
# --------------------------------------------------------------------------- #
def _family_note(w, family, sheet, charts, templates):
    name = sheet.name
    hdr_row, hdr_cols = sheet.header_row()
    first, last = sheet.data_rows(hdr_row)

    w("## Building %s, on `%s`" % (FAMILY_HEAD[family], name))
    w("")
    w("Part 1 applies unchanged. What differs:")
    w("")

    if family == "table":
        w("- **The grid is the deliverable.** There is no chart zone; the "
          "table itself is what prints, from row %s down. Set the column "
          "widths by hand and the row heights to a fixed pitch, because the "
          "drawn columns are positioned against those rows."
          % (hdr_row or "?"))
        w("- **Scenario notation lives in the borders.** A column header over "
          "an actual gets a bottom border in the scenario's own weight and "
          "colour; the double rule that says *plan* is Excel's own "
          "**xlDouble** border style - the outlined fill seen edge-on. Nothing "
          "here needs a drawn shape.")
        w("- **A drawn column is an ordinary variance chart** sized to the row "
          "pitch and positioned over the block it belongs to, with no axis, no "
          "gridlines and a transparent chart area. It is the tier stack of "
          "Part 2, one column wide.")
        w("- **The threshold is a cell.** Conditional formatting tests against "
          "the same cell the footnote displays, so the note and the colouring "
          "cannot disagree. IBCS requires the threshold to be stated wherever "
          "red is applied.")
        w("- **Subtotals are formulas**, never typed:")
        w("")
        w.extend(F.quote(F.find_example(templates, "subtotal")))
    elif family == "structure":
        w("- **One data block per panel**, side by side, and every panel on "
          "**one shared scale** - which is the entire point of the template. "
          "Work out the tallest column across all panels, round it up, and "
          "give every panel that maximum. Do not let Excel scale each one.")
        w("- **A stacked category is not a stacked scenario.** Where the bands "
          "are business areas, every band is an actual, so the fill cannot "
          "carry the scenario any more. The *column* carries it - a plan "
          "column is outlined - and the bands take a light-to-dark ramp.")
        w("- **Insert as a Stacked Column (or Bar)**, one series per band, "
          "bottom band first.")
        w("- **Small bands get no label.** Below about 3%% of the axis a "
          "number does not fit, and a label that overlaps its neighbour is "
          "worse than an absent one. That test belongs in the label formula, "
          "not in your judgement:")
        w("")
        w.extend(F.quote(F.find_example(templates, "text")))
    elif family == "line":
        w("- **The data runs in rows, not columns** - one row per series "
          "across the periods. Everything in Part 1 still applies, transposed.")
        w("- **A line must show its markers.** The line between two points is "
          "a connector, not data - there is no value at half past March - so "
          "the marker is what says measured, planned or expected. That is how "
          "one line changes from actual to forecast part way along, and it is "
          "why you cannot use a plain line chart here.")
        w("- **It is a combo chart**: columns for the monthly tier, lines for "
          "the cumulative. **Chart Design -> Change Chart Type -> Combo**, and "
          "set the type per series.")
        w("- **A flow and a level cannot share an x position.** A stock is "
          "drawn at period *boundaries* - an opening balance and one closing "
          "per period - while the movements that produce it sit inside the "
          "periods they belong to. Give them the same x and neither means "
          "anything.")
    elif family == "XY":
        w("- **This is the one family that shows its value axis**, so the "
          "scale block of Part 1 step 5 **cannot be used**. Divide by a span "
          "here and the axis prints `0.25` where your data says `29.16%`. Fit "
          "the axis to the data once instead, and accept that it will not "
          "follow an edit.")
        w("- **Insert -> Scatter, or Bubble.** Each category that needs its own "
          "colour is its own series, because a series takes one colour.")
        w("- **A mark's area carries the value; its radius never does.** "
          "Bubble area goes as the value, so the radius goes as its square "
          "root. Excel's *Represent bubble size as: Area* does this for you - "
          "check it is set, because *Width* is also on that menu and it "
          "overstates a large value by the square of the ratio.")
        w("- **A colour that carries a category needs a key; a scenario never "
          "does.** Solid dark means measured on every IBCS page ever printed. "
          "A product line means nothing outside its own chart, so it takes an "
          "accent colour - deliberately not one of the scenario greys - and a "
          "legend. The presence of the key is itself the signal that the "
          "colour is not notation.")
        w("- **Where marks overlap, paint order is notation.** The mark the "
          "reader is being asked about goes last, even where it is the smaller "
          "one. Reorder with **Select Data -> the up/down arrows**.")
    elif family == "tree":
        w("- **Each box is its own small chart**, and the arithmetic between "
          "them is text boxes and lines. There is no tree chart type and there "
          "is no way to fake one.")
        w("- **Boxes that share a unit share a scale.** Return, net sales and "
          "capital are all kEUR, so all three get the same points-per-unit - "
          "otherwise the tree invites a comparison it does not support.")
        w("- **The connectors state a calculation.** `ROS x Turnover = ROI` is "
          "written on the page, so the reader can check the tree against "
          "itself.")
        w("- Group the whole thing (**select all -> right-click -> Group**) "
          "once it is right, or moving it will take it apart.")
    elif family == "panel":
        w("- **Do not build this by copying one chart R x C times.** That is "
          "the usual advice and it leaves you maintaining a dozen charts whose "
          "axes drift apart the first time the data changes.")
        w("- The whole grid is **one chart object**, with the value axis "
          "divided into *bands* - one per grid row - and each panel's data "
          "arithmetically squeezed into its band:")
        w("")
        w("```")
        w("plotted = band + (value - vmin) / (vmax - vmin) * band_frac")
        w("```")
        w("")
        w("- Everything that is not data - the dividers, the panel titles, the "
          "tick labels - is drawn by extra XY scatter series pretending to be "
          "chart furniture, on the **secondary** axes.")
        w("- This one is genuinely hard to do by hand. The companion "
          "`panel-charts` skill builds it, and its `references/manual-steps.md` "
          "is a walkthrough of its own; that is the right document for this "
          "family rather than a paragraph here.")
    w("")

    if charts:
        w("**The charts on this sheet, as built:**")
        w("")
        w("| Chart | Type | Left | Top | Width | Height | Series |")
        w("|---|---|---|---|---|---|---|")
        for c in charts:
            w("| `%s` | %s | %g | %g | %g | %g | %d |"
              % (c["name"], _ctype(c), c.get("left", 0), c.get("top", 0),
                 c.get("width", 0), c.get("height", 0),
                 len(c.get("series", []))))
        w("")


# --------------------------------------------------------------------------- #
# Part 3 - sheet by sheet
# --------------------------------------------------------------------------- #
def _sheet_section(w, s, charts):
    w("### %s" % s.name)
    w("")
    message = s.value(6, 2)
    if isinstance(message, str) and message:
        w("> %s" % message)
        w("")
    w("**%s.**" % FAMILY_HEAD[F.family_of(s.name)].capitalize())
    w("")
    w("| | |")
    w("|---|---|")
    w("| Cells used | `%s` |" % s.ws.dimensions)
    w("| Print area | `%s` |" % (F.plain(s.ws.print_area) or "-"))
    if s.hidden:
        w("| Hidden scale columns | %s |"
          % ", ".join("`%s`" % h for h in s.hidden))
    w("| Formula cells | %d, in %d shapes |" % (s.n_formulas, len(s.formulas)))
    w("| Typed cells | %d |" % len(s.typed))
    w("")
    if s.typed:
        w("**Typed values** - %s. Everything else is a formula over these. The "
          "cells a reader is meant to edit are shaded; a sheet may also park a "
          "constant the engine needs in that range, so go by the shading."
          % F.ranges(s.typed))
        w("")

    groups = s.groups()
    named = [(g, s.header_for(g[1][0][0], g[1][0][1])) for g in groups]
    shown = [(g, h) for g, h in named if len(g[1]) > 1 or h]
    rest = [g for g, h in named if not (len(g[1]) > 1 or h)]
    w("**Every formula on the sheet.** Cells that say the same thing about "
      "their own position are one row here, with the range they cover and the "
      "first of them written out.")
    w("")
    w("| Cells | Named | Example | Formula |")
    w("|---|---|---|---|")
    for (_shape, cells), head in shown[:F.TABLE_CAP]:
        row, col, formula = cells[0]
        heads, seen = [], set()
        for c in sorted({c for _r, c, _f in cells}):
            nm = s.header_for(row, c)
            if nm and nm not in seen:
                seen.add(nm)
                heads.append(nm)
        w("| %s | %s | `%s` | `%s` |"
          % (F.ranges([(r, c) for r, c, _f in cells]),
             ", ".join(heads[:4]) or head or "-", F.addr(row, col),
             formula.replace("|", F.BAR)))
    if len(shown) > F.TABLE_CAP:
        w("")
        w("*%d further named shapes are not listed - this sheet has %d in all.*"
          % (len(shown) - F.TABLE_CAP, len(shown)))
    if rest:
        w("")
        w("%d unnamed one-off formula%s: %s."
          % (len(rest), "" if len(rest) == 1 else "s",
             ", ".join("`%s`" % F.addr(c[1][0][0], c[1][0][1])
                       for c in rest[:12])
             + (", ..." if len(rest) > 12 else "")))
    w("")

    if charts:
        w("**Charts**, in points from the top left of the sheet.")
        w("")
        w("| Name | Type | Left | Top | Width | Height | Series | Axis |")
        w("|---|---|---|---|---|---|---|---|")
        for c in charts:
            axis = ("%s to %s" % (_fmt(c["min"]), _fmt(c["max"]))
                    if "min" in c else "-")
            w("| `%s` | %s | %g | %g | %g | %g | %d | %s |"
              % (c["name"], _ctype(c), c.get("left", 0), c.get("top", 0),
                 c.get("width", 0), c.get("height", 0),
                 len(c.get("series", [])), axis))
        w("")
        w("| Chart | Series | Fill | Marker | Stem | Labelled |")
        w("|---|---|---|---|---|---|")
        for c in charts:
            for ser in c.get("series", []):
                # Whenever the points are not all painted the same - a
                # marker series, or a series with the forecast months hatched -
                # the series-level fill is not the whole truth.
                groups = ser.get("paint") or []
                paint = (_paint(ser)
                         if len(groups) > 1
                         or ser.get("marker") not in (None, "None")
                         else ser.get("fill", "-"))
                w("| `%s` | `%s` | %s | %s | %s | %s |"
                  % (c["name"], ser.get("name", "?"), paint,
                     ser.get("marker", "-"),
                     ("%s %s" % (ser.get("stem"), ser.get("stem_colour", "")))
                     .strip() if ser.get("stem") else "-",
                     "yes" if ser.get("labels") else ""))
        w("")


# --------------------------------------------------------------------------- #
class _Writer:
    """A callable that appends a line, and also splices a block of them.

    The section writers want both - `w("a line")` reads better than
    `out.append(...)` a hundred times over, and a quoted formula arrives as a
    list. A plain bound method cannot carry the second, so this does.
    """

    def __init__(self, out):
        self.out = out

    def __call__(self, line):
        self.out.append(line)

    def extend(self, lines):
        self.out.extend(lines)


def document(book, sheets, charts, simple=False):
    out = []
    w = _Writer(out)
    templates = [s for s in sheets if s.name != "Read me"]

    _front(w, book, templates, charts, simple)
    _skeleton(w, templates, charts)

    w("# Part 2 - Worked builds")
    w("")
    w("One per family present in this workbook. The tier stack is done in "
      "full; the rest are written as what differs from it.")
    w("")

    seen = set()
    by_family = {}
    for s in templates:
        by_family.setdefault(F.family_of(s.name), []).append(s)
    for family in FAMILY_ORDER:
        members = by_family.get(family)
        if not members:
            continue
        pick = next((m for m in members if m.name in F.TEACHING_ORDER),
                    members[0])
        seen.add(pick.name)
        if family == "tier stack":
            _tier_walkthrough(w, pick, charts.get(pick.name, []), templates)
        else:
            _family_note(w, family, pick, charts.get(pick.name, []), templates)
        w("---")
        w("")

    w("# Part 3 - Sheet by sheet")
    w("")
    w("What is on each sheet, for rebuilding one template rather than learning "
      "the method.")
    w("")
    for s in templates:
        _sheet_section(w, s, charts.get(s.name, []))

    w("---")
    w("")
    w("*Generated by `ibcs_doc.py` from the workbook itself. Every formula, "
      "colour and measurement above was read back out of it.*")
    return NL.join(out) + NL
