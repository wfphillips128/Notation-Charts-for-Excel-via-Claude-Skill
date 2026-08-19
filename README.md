# Excel notation charts

Management-report charts and tables in **IBCS® notation**, built as native Excel
chart objects with live formulas — and as self-contained SVG — from one shared
data layer.

All **17 IBCS® templates** are recreated end to end: C01–C13 charts and T01–T04
tables. Paste your own figures into a sheet's input block and the chart, its
captions, its variance colours and its axes all follow.

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
| [`ibcs/`](ibcs/) | The skill: one data layer, two renderers, and the reasoning behind every choice |
| [`workbooks/`](workbooks/) | Two finished `.xlsx` files and a generated build document for each |

### The workbooks

| File | What it holds |
|---|---|
| `IBCS Charts - Complex Versions.xlsx` | all 17 templates, every tier |
| `IBCS Charts - Simple Variants Only.xlsx` | the seven templates with a useful base-tier form, drawn with that tier alone |

Each opens on a **Read me**, then one sheet per template in template-id order.
Every sheet prints on **one page**.

Beside each workbook is a `.md` of the same name. That document is **generated
from the workbook** — it opens the `.xlsx` and reads the formulas back out of
it, so it cannot describe a formula the file does not contain. It is the manual
build: what to type, where, and which of it is a trap.

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

### The Excel and the SVG cannot disagree

They are two renderers over one data layer. Geometry is solved in SVG first,
where working out where a pin head goes is arithmetic, and then ported to
Excel, where it is a fight. Fidelity is measured rather than judged — C12A's
eighty bars land within 4px of the published reference, C09C's 149 markers 144
of them exact to the centre pixel, and so on for all seventeen.

---

## Installing the skill

This is a [Claude Code](https://claude.com/claude-code) skill. Copy the folder
into your skills directory and restart Claude Code:

```bash
git clone https://github.com/wfphillips128/excel-notation-charts.git
cp -r excel-notation-charts/ibcs ~/.claude/skills/          # macOS / Linux
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
python .../ibcs/scripts/ibcs_svg.py --template C04A

# The Excel workbook, and the build document read back out of it
python .../ibcs/scripts/ibcs_excel.py --template C03A,C04A \
    --out book.xlsx --doc book.md

# The gates
python .../ibcs/scripts/ibcs_data.py          # 127 tie-outs
python .../ibcs/scripts/test_responsive.py    # the workbook is still live
python .../ibcs/scripts/test_rescale.py       # it follows other numbers
python .../ibcs/scripts/ibcs_doc.py --xlsx book.xlsx --out book.md --check
```

Output goes to `./build` by default — the current directory, not next to the
scripts. Set `IBCS_BUILD` to put it elsewhere.

## The one dependency worth explaining

`C13D` draws fifteen locations as small multiples in a **single native chart
object**, not fifteen copies of one chart whose axes drift apart the first time
a figure changes. Excel has no panel chart type, so that grid is built by a
companion skill, `panel-charts`, which is not in this repository yet.

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
and both say so plainly if the directory is absent. Nothing else needs them;
everything this project builds, it builds from `ibcs_data.py`.

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
