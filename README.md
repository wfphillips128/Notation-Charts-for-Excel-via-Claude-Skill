# Notation charts for Excel

Management-report charts and tables in **IBCS® notation**, built as native Excel
chart objects with live formulas — and as self-contained SVG — from one shared
data layer.

All **17 IBCS® templates** are recreated end to end: C01–C13 charts and T01–T04
tables. Paste your own figures into a sheet's input block and the chart, its
captions, its variance colours and its axes all follow.

They are drawn twice, from two datasets. Once from the IBCS® Institute's own
published figures, which is what makes this a fidelity recreation — and once
from **The Progressive Corporation's SEC filings**, which is what makes it a
workbook you can paste over without first working out whose numbers you are
replacing.

> ### Requirements, up front
>
> The Excel half of this project drives **Excel itself** through COM. It needs
> **Windows**, **Excel 2016 or later**, and **`pywin32`**. There is no
> cross-platform path and there is not going to be one — the whole point is a
> workbook a colleague can open, edit and print, not a picture of one.
>
> The **SVG renderer is pure Python** and runs anywhere. If you only want the
> pictures, that half has no dependencies at all.
>
> ```bash
> pip install pywin32 openpyxl        # Excel side
> ```

---

## What is actually in here

| | |
|---|---|
| [`ibcs-excel-charts/`](ibcs-excel-charts/) | The skill: two data layers, two renderers, and the reasoning behind every choice |
| [`workbooks/`](workbooks/) | Four finished `.xlsx` files — two datasets, two forms each |
| [`guides/`](guides/) | How to build each of them by hand in Excel |
| [`docs/`](docs/) | Design notes, including the one that planned the second dataset |

### The workbooks

| File | Data | What it holds |
|---|---|---|
| `IBCS Charts - Complex Versions.xlsx` | Institute | all 17 templates, every tier |
| `IBCS Charts - Simple Variants Only.xlsx` | Institute | the seven templates with a useful base-tier form, drawn with that tier alone |
| `IBCS Charts-Progressive - Complex Versions.xlsx` | Progressive | all 17, every tier |
| `IBCS Charts-Progressive - Simple Variants Only.xlsx` | Progressive | the same seven, that tier alone |

**Which pair you want.** The Institute pair *is* the recreation — its cells hold
the published figures, and it exists to be measured against its source. The
Progressive pair is the one to paste over: every figure in it belongs to a real
company, so none of them is a number you have to identify before you can replace
it. See [The second dataset](#the-second-dataset).

Each opens on a **Read me**, then one sheet per template in template-id order;
the Progressive pair carries a **Sources** sheet as well. Every sheet prints on
**one page**.

In [`guides/`](guides/) is an **Excel Guide** for each — how to build that workbook by hand
in Excel, with nothing but the menus. Part 1 is the sheet skeleton every template
shares, Part 2 a worked build per chart family, Part 3 a per-sheet reference.

Those guides are **generated from the workbooks**. Every range, formula, hex
colour, gap width and axis bound in them was read back out of the `.xlsx` after
it was written — including how each individual *point* is painted, which is
where the forecast hatch and a pin head's scenario live. A guide that tells you
to set a fill to `#404040` is only worth having if that value came out of the
file.

---

## Why this is not a template pack

IBCS® is a **notation**, not a style guide. Its claim is narrow and strong: the
same meaning must always take the same visual form, so a reader who has learned
the notation once can read any report without a legend. A solid dark bar is
something that happened; an outlined one is something planned; a hatched one is
a forecast. Red means an adverse impact, which is not the same as a negative
number.

Most of what goes wrong in a management report is not ugliness — it is
ambiguity. A forecast drawn like an actual. A cost overrun coloured green
because the number was positive. A variance tier with nothing to say what it is
a variance *from*.

Four things here follow from taking that seriously, and they are the reason
this is code rather than a folder of `.xltx` files.

### The charts follow *your* numbers

Excel will not bind an axis bound to a formula — an axis maximum is a number
you type. So a chart built for one set of figures clips or shrinks the moment
somebody pastes their own over the top, which is what makes most Excel chart
templates single-use.

These sheets invert the problem: instead of scaling the axis to the data, they
**scale the data to a fixed axis**. A hidden span cell per group of tiers
sharing a unit, a scaled copy of every drawn column, and axis bounds divided by
the same span. Dividing values and bounds by one number is a visual no-op at
today's figures and follows the data at any others — and because zero divided
by anything is zero, the zero line, the reference rules and every caption
positioned from plot geometry stay exactly where they were.

`test_rescale.py` pastes figures a hundred times larger into each sheet and
asserts that nothing plots outside its axis and that tiers sharing a unit still
agree.

Two limits, stated rather than hidden. **C09C and C10D are not normalised** —
they are the only charts that show a value axis, and dividing by a span would
print `0.25` where the data says `29.16%`; they are fitted at build time
instead. And **C12A's relative tier clips rather than scales**, because its
declared range is narrower than its data's span, so the largest value would
clip on any data at all. It draws the clip and puts the overflow in the label.

### Every word a chart shows comes from a cell

The sheet's text block is six typed inputs — entity, measure, unit, period,
reference scenario, message — and every line the reader sees is a formula over
them, with the chart's text boxes *linked* to those cells. Retype the unit once
and the subject line and every tier caption follow. Nothing a chart says is
typed into the chart.

### Nothing renders on a broken tie

`check_ties()` proves the arithmetic against the labels the source actually
printed — 127 checks — and both renderers call it before drawing anything.
A total with components is a formula, never an input: a total that does not
follow its parts is the same lie a static variance colour is.

That gate is **vacuous on any other data, and it fails open**. Its per-template
checks read the recreation's own module globals, so handed a template built from
somebody else's figures it re-proves the Institute's numbers against themselves
and passes. The second dataset therefore brings its own — `test_alt_ties.py`,
**211 checks**, every one of them reading the template it was given — and it
runs before the workbook is built rather than after.

### The Excel and the SVG cannot disagree

They are two renderers over one data layer. Geometry is solved in SVG first,
where working out where a pin head goes is arithmetic, and then ported to
Excel, where it is a fight. Fidelity is measured rather than judged — C12A's
eighty bars land within 4px of the published reference, C09C's 149 markers 144
of them exact to the centre pixel, and so on for all seventeen.

---

## The second dataset

A fidelity recreation is a poor file to paste your own numbers into. Its cells
hold the Institute's published figures, and you cannot tell which of them are
Furniture Inc.'s without checking each one. That is the point of a recreation
and the problem with using it as a template.

So the seventeen are drawn a second time, from **The Progressive Corporation
(PGR)** — a real insurer's real, public figures.

**Nothing is transcribed by hand.** `fetch_pgr.py` harvests the filings from SEC
EDGAR, `pgr_parse.py` reads them, and `ibcs_data_alt.py` reads what those write
into [`datasets/pgr/`](ibcs-excel-charts/datasets/pgr/). That is what lets the
figures be checked rather than trusted. `test_pgr_reconcile.py` runs **605
checks against Progressive's own totals**: twelve monthly premium figures from
twelve separate news releases summing to the annual number in the audited 10-K,
every release's prior-year column agreeing with the release of twelve months
before, segments adding to companywide. A reconciliation against the source is a
stronger claim than internal consistency, and it is the one thing invented
figures could never have offered.

**Eleven of the seventeen templates want a plan or a forecast, and no public
company files one.** Those are constructed from Progressive's own long-stated
underwriting target — a **96 combined ratio**, which is a 4% margin — so plan
underwriting profit is `net premiums earned × 4 / 100`. The rule is printed on
each sheet that uses it.

**A constructed figure is not allowed to look reported.** Each template declares
where its scenarios came from, and the workbook shades an *assumed* scenario's
typed cells differently from a reported one's, with a **Sources** sheet saying
what was assumed and why. `test_provenance.py` checks both directions — that the
marking marks, and that a template declaring nothing comes out exactly as it did
before the facility existed. A feature that quietly restyles already-published
work is a defect however good the feature is.

**Neither dataset can leak into the other.** Three templates used to reach past
the template they were handed and read module globals — C08H's opening balance,
C09C's iso-curves, C13D's panel figures — which fails *silently*: a global left
behind draws one company's numbers on a sheet titled with another's.
`test_dataset_isolation.py` asserts the property rather than the fix. Every name
a renderer reads out of `ibcs_data` must be a type, a helper, a registry or a
`Template`; anything else is data, and data arrives through the argument.

---

## Installing the skill

This is a [Claude Code](https://claude.com/claude-code) skill. Copy the folder
into your skills directory and restart Claude Code:

```bash
git clone https://github.com/wfphillips128/Notation-Charts-for-Excel-via-Claude-Skill.git
cp -r Notation-Charts-for-Excel-via-Claude-Skill/ibcs-excel-charts ~/.claude/skills/
```

On Windows the destination is `%USERPROFILE%\.claude\skills\`.

Or drop it into a single project's `.claude/skills/` instead, if you would
rather it did not load everywhere.

You do not need Claude Code to use it. The scripts are ordinary Python and the
workbooks are ordinary workbooks.

## Using it directly

```bash
cd wherever-you-want-the-output

# One template to SVG - no dependencies beyond Python
python .../ibcs-excel-charts/scripts/ibcs_svg.py --template C04A

# The Excel workbook, and the build document read back out of it
python .../ibcs-excel-charts/scripts/ibcs_excel.py --template C03A,C04A \
    --out book.xlsx --doc "book - Excel Guide.md"

# The same templates on the Progressive data - its own entry point, because
# ibcs_excel's CLI resolves templates from ibcs_data
python .../ibcs-excel-charts/scripts/build_alt.py pgr.xlsx \
    --doc "pgr - Excel Guide.md"

# The gates - the recreation
python .../ibcs-excel-charts/scripts/ibcs_data.py             # 127 tie-outs
python .../ibcs-excel-charts/scripts/test_responsive.py       # the workbook is still live
python .../ibcs-excel-charts/scripts/test_rescale.py          # it follows other numbers
python .../ibcs-excel-charts/scripts/ibcs_doc.py --xlsx book.xlsx --check

# The gates - the second dataset
python .../ibcs-excel-charts/scripts/test_alt_ties.py         # 211 tie-outs
python .../ibcs-excel-charts/scripts/test_pgr_reconcile.py    # 605 checks against PGR's own totals
python .../ibcs-excel-charts/scripts/test_dataset_isolation.py  # no renderer reaches past its template
python .../ibcs-excel-charts/scripts/test_provenance.py       # constructed figures are shaded as such
```

Output goes to `./build` by default — the current directory, not next to the
scripts. Set `IBCS_BUILD` to put it elsewhere.

## The one dependency worth explaining

`C13D` draws fifteen locations as small multiples in a **single native chart
object**, not fifteen copies of one chart whose axes drift apart the first time
a figure changes. Excel has no panel chart type, so that grid is built by a
companion skill,
[`panel-charts`](https://github.com/wfphillips128/Panel-Charts-for-Excel-via-Claude-Skill),
which lives in its own repository.

The sheet in the shipped workbook is complete. Rebuilding *that one sheet* from
source needs the companion: install it beside this skill, or set `PANEL_CHARTS`
to its `scripts` directory. Without it, C13D fails with a message saying so and
the other sixteen templates build normally.

---

## Attribution and disclaimer

IBCS® and International Business Communication Standards are trademarks of the
IBCS Association. This project is an **independent recreation, made for study,
from the published IBCS® templates**. It is not an IBCS Institute publication
and is not endorsed by or affiliated with the IBCS Association.

**The IBCS® reference renders are not distributed here.** They are the
Institute's own images and were used only to measure a recreation against its
source. Two scripts want them — `compare_render.py` and `extract_palette.py` —
and both say so plainly if the directory is absent. Nothing else needs them: the
recreation is built from `ibcs_data.py` and the Progressive workbooks from
`ibcs_data_alt.py`, and neither renderer reads an image.

**Not affiliated with The Progressive Corporation.** The Progressive workbooks
are an independent presentation of figures that company published in its SEC
filings and investor-relations releases. They are not that company's reporting,
and the plan and forecast scenarios in them were never published by it — they
are constructed here, on a basis stated on the sheet and on the Sources sheet.

**IBCS® prescribes no colour codes.** Rule UN 4.1 says so outright. The palette
in `assets/ibcs-palette.json` is a house choice with its provenance recorded,
not a mandated one, and no hex value here should be presented as a requirement
of the standard.

The standard itself is published at [ibcs.com](https://www.ibcs.com/). If you
are producing management reports for a living, read it.

## Licence

MIT — see [LICENSE](LICENSE). That covers this project's code and workbooks; it
does not and cannot grant anything in respect of the IBCS® standard or the
Institute's own materials.
