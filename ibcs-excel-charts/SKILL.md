---
name: ibcs-excel-charts
description: Build charts and tables in IBCS® notation (International Business Communication Standards) - the "what means the same must look the same" standard behind scenario notation (AC/PY/PL/BU/FC), variance tiers, message-driven titles, and the 17 IBCS® templates. Produces Excel charts via win32com and self-contained SVG from one shared data layer. Use when the user mentions IBCS, SUCCESS rules, scenario notation, variance charts with pins, multi-tier charts, "what means the same must look the same", Hichert, Zebra BI-style reporting, or asks for a management report chart that follows a formal notation standard.
---

# IBCS® notation

IBCS® is a notation standard, not a style guide. Its claim is narrow and strong:
the same meaning must always get the same visual form, so a reader who has
learned the notation once can read any report without a legend.

Most of what goes wrong in practice is not ugliness, it is ambiguity - a
forecast drawn like an actual, a cost overrun coloured green because the number
was positive, a variance tier with no indication of what it is a variance from.

## The rules that carry the most weight

**Fill encodes how a number came to exist**, never what it measures (UN 3.2):

| Scenario | Fill | Means |
|---|---|---|
| AC | solid dark | actual - measured |
| PY | solid light | earlier actual - measured |
| PL / BU | outlined, no fill | plan or budget - fictitious |
| FC | hatched, 45° ascending | forecast - expected |

**Variance colour encodes impact, not sign.** A cost 40 over plan is a positive
number and a bad outcome, so it is red. Call
`ibcs_style.variance_colour(value, higher_is_better=False)` for cost and expense
measures. Genuinely ambiguous measures - headcount, inventory, mix - are blue,
not forced into good/bad.

**The variance axis carries the reference scenario**, which is why a correct
variance tier needs no legend: solid light for a ΔPY comparison, a double rule
for ΔPL or ΔBU.

**Absolute variances are columns; relative variances are pins** - a thin stem
with a head marker that carries the minuend's scenario fill. The head is
load-bearing: it is what says whether a relative variance was measured or
forecast.

**The title is a message, not a label.** "Contribution in kEUR" is the subject.
The title is the sentence stating what the reader should conclude.

**A period name can carry an analysis.** A leading underscore means year to date
("_Jun 2026"), a trailing one means year to go ("Jun 2026_"), and a leading tilde
means a moving annual analysis - MAT for a total, MAA for an average (UN 4.2).
This is why the second block of a T01 table is headed `_November` rather than
"YTD", and it is notation: the reader is told what a column cumulates without a
word being spent on it.

**A mark's area carries a value; its radius never does.** Bubble area goes as
the value, so the radius goes as its square root. Making the radius proportional
instead overstates a large value by the square of the ratio - on C10 that would
draw a 32 against a 1.1 twenty-nine times too big. The size legend states what
the area counts, because on a chart plotting three measures no single unit can
head the page.

**A colour that carries a category needs a key; a scenario never does.** Solid
dark means measured on every IBCS® page ever printed, so scenario fills are
self-explaining. A product line means nothing outside its own chart, so it takes
an accent colour - deliberately not one of the scenario greys - and a legend.
The presence of the key is itself the signal that the colour is not notation.

**Where marks overlap, paint order is notation.** The mark the reader is being
asked about goes last. On a portfolio chart that is the actual, even where the
actual is the smaller mark and size-ordering would bury it under the prior year.

**IBCS® prescribes no colour codes.** Rule UN 4.1 (p52) says so outright. Any
IBCS® palette, including this skill's, is a house choice. Do not present hex
values found online as mandated.

## How to build

One data layer, two renderers, so the Excel workbook and the web output cannot
disagree:

```
scripts/ibcs_data.py     template data, structure, and check_ties()
scripts/ibcs_style.py    semantic constants - ask for "the fill for a forecast"
scripts/ibcs_layout.py   per-template sheet layout: data zone, chart zone, tiers
scripts/ibcs_svg.py      SVG renderer (also the geometry reference)
scripts/ibcs_excel.py    win32com renderer, reproducing the SVG geometry
scripts/ibcs_doc.py      reads a built workbook back: formulas, and how every
                         chart and every point is actually painted
scripts/ibcs_guide.py    the by-hand build guide written around those facts
scripts/ibcs_paths.py    where output goes, and where the companion panel-charts
                         skill is - found rather than assumed
scripts/test_responsive.py  proves the workbook is still live after an edit
scripts/test_rescale.py     proves a sheet follows figures 100x larger
scripts/test_dataset_isolation.py  proves neither renderer reads a figure from
                            the data module - everything drawn arrives on the
                            Template it was handed
scripts/test_provenance.py  proves a constructed figure is marked as one, and
                            that the marking stays silent on data that has
                            none
scripts/compare_render.py   measures a render against the IBCS(R) original;
                            wants $IBCS_TEMPLATE_REFS, which is not shipped
scripts/extract_palette.py  reads fills out of those same references
assets/ibcs-palette.json the house palette, with provenance
```

A second dataset adds six scripts and a `datasets/` tree to that list - see
**Drawing other data** below.

**A table is the same data layer, a second layout engine.** A column of a table
is a `Tier` - a keyed set of values, optionally measured against a scenario - and
a hierarchy of rows with subtotals is what `Row` already described for a P&L, so
nothing in the data layer changed. What a table adds is a grid: declared column
widths, rows at a pitch, and the header rules that carry the scenario notation.
A column that is *drawn* rather than printed is named in `Template.panel_tiers`
and hangs the ordinary variance primitives inside the same row band.

**A worksheet has two zones and nothing straddles them**: a data zone of declared
width on the left, the charts to its right, with the chart zone's left edge
measured from where the data ends. Never `AutoFit` a data block a chart is
positioned against - one long string moves the boundary and the charts end up
drawn over the source data. `verify()` asserts the separation.

**Solve geometry in SVG first, then port to Excel.** Working out where a pin head
goes is arithmetic in SVG and a fight in Excel. See `references/decisions.md`
for the Excel traps that cost the most time - the GapWidth ceiling, the error-bar
pin technique, and the plot-area assignment order.

**Every word a chart shows comes from a cell.** The sheet's text block is six
typed inputs - entity, measure, unit, period, reference scenario, message - and
the lines the reader sees are formulas over them, with the chart's text boxes
linked to those cells. Retype the unit once and the subject line and every tier
caption follow. Nothing a chart says may be typed into the chart.

**A flow and a level cannot share an x position.** A stock is drawn at period
*boundaries* - an opening balance and one closing per period - while the
movements that produce it are drawn inside the periods they belong to. Give them
the same x and neither means anything. And where both appear, they share a zero:
the increase columns rise into the stock they add to.

**A line must show its markers, and the marker carries the scenario.** The line
between two points is a connector, not data - there is no value at half past
March - so IBCS® requires markers, and since a line has no fill they are what says
measured, planned or expected. That is how one line changes from actual to
forecast part way along.

**A stack of categories is not a stack of scenarios.** Where the bands are
business areas or channels, every one of them is an actual, so the fill cannot
carry the scenario any more - the *column* does, and the bands take a structure
ramp. That is why they are `Segment`s rather than `Series`, and why a plan
column is outlined rather than any band inside it being shaded differently.

**A total is data.** Templates that show one hold it in `Template.summary_rows`,
which both renderers read, and a total with components is a formula rather than
an input - a total that does not follow its parts is the same lie a static
variance colour is. Where that makes the workbook disagree with the reference by
a rounding, the SVG prints what the original printed and the workbook prints
what its own data says.

**Impact is per row, not per column.** Within one ΔPL column of a statement the
favourable bars sit on both sides of the axis, because a cost line up is adverse
and a cost line down is favourable. Split a variance series by *impact* - what
`higher_is_better_at` says - never by the sign of the number, or every cost
overrun comes out green.

**The charts follow the reader's numbers, not IBCS®'s.** Excel will not bind an
axis bound to a formula - an axis maximum is a number you type. So instead of
scaling the axis to the data, the sheets **scale the data to a fixed axis**: a
hidden **span cell** per group of tiers sharing a unit, a **scaled copy** of
every drawn column (`= raw / span`), and axis bounds that are the declared
bounds divided by the same span. Dividing values and bounds by one number is
visually a no-op at today's figures and follows the data at any others, and
because zero divided by anything is zero the zero line, the reference rules and
every caption positioned from plot geometry stay put. `scripts/test_rescale.py`
pastes figures a hundred times larger into each sheet and asserts nothing plots
outside its axis.

**The claim above is the one this codebase keeps failing to honour, so it is
worth knowing where it breaks.** Dividing the *values* by a span is only half of
it: the *bounds* have to be in the same units, and for a long time they were not.
The layout declared C03A's axis as `230` - kEUR, measured against the Institute's
figures - and divided it by the live span. On a company forty-seven times larger
that put the axis at 0.023 while every normalised bar sat near 1.0, Excel clamped
them all to the frame, and the chart drew flat while the data varied fourfold.
Nothing caught it: the geometry verification passed, `test_rescale` passed
(multiplying inputs scales the span too), the tie-outs passed. Only the picture
was wrong.

Seven separate defects turned out to share that shape, in **three flavours**, and
only the first yields to "declare it as a fraction of the span":

| | Examples |
|---|---|
| **Magnitude** | axis bounds in kEUR; `LineLayout.maximum = 24` above an inventory of 22 tons; a number format of `0.0` printing `17258.0`; a label-fit floor as a fixed share of the axis; a tolerance of ±1.5 on figures a thousand times larger; `PanelGeometry.scale` in pixels per kEUR |
| **Shape** | padding and asymmetry that encode *how far* the reference's bridge travels, or *which way* its variances point. These survive conversion to fractions and are still wrong: a window sized for a walk of 14% of its base puts the axis floor below zero when the walk is 228% |
| **Label** | `'2025 PL'`, a summary row looked up by name - a fact about one year of one company. Raised `KeyError` on the first sheet built from anything else |

So: **a constant in the layout should describe the template, never the figures.**
Where it cannot, the Template carries its own - `tier_bounds`, `panel_geometry`,
`TreeSpec.scale_px`, `StructurePanel.number_format`, `Row.ratio_of` - and the
layout's value is the fallback. And when a constant moves onto the Template,
*every reader of it has to move too*: `test_rescale` had to learn the same lesson
three times, because it was asserting the mechanism (`axis x span == the declared
bound`) rather than the property.

The lesson crosses the skill boundary. `panel-charts` exposes its axis floor as
`PanelSpec.floor_pad` for the same reason, and `verify_panel_grid` reads it off
the spec rather than retyping the number - two copies of one constant is one
drift, whichever side of a dependency they sit on.

`verify_axis_containment` runs inside `build()` for every sheet of every dataset
and reads each chart's axis and every drawn point back out of Excel. It is
falsifiable both ways - it fires on 24 of 24 points on the original bug, and is
silent on correct sheets - and it is the only check that looks at the picture.

It looks at *geometry*, though, and there is one thing it cannot see. Nothing
here measures a rendered label's box, so a value label printing through a rule
survives every gate: the digits are struck through, not missing, and everything
that counts cells passes. That is what `--export-dir` is for on both builders.
**On the panel grid, export the PNG and look at it.**

Two limits, stated rather than hidden. **C09C and C10D are not normalised** -
they are the only charts that show a value axis, so dividing by a span would
print `0.25` where the data says `29.16%`; they are fitted at build time
instead, correct for the data present when the sheet is built but not live on
edit. And **C12A's relative tier clips rather than scales**: its declared range
is narrower than its data's span, so the largest value would clip on any data
at all. It draws the clip and puts the overflow in the label - `+983`.

**Never render on a broken tie.** `check_ties()` proves the arithmetic against
the labels the source actually printed, and both renderers call it before
drawing.

```bash
python scripts/ibcs_data.py                        # tie-outs and the derived table
python scripts/ibcs_svg.py --template C04A         # render - always before
                                                   # comparing: compare_render
                                                   # reads the file, it does not
                                                   # call the renderer
python scripts/compare_render.py C04A              # diff against the IBCS® original
python scripts/ibcs_excel.py --template C03A,C04A --out book.xlsx        --doc "book - Excel Guide.md"
python scripts/test_responsive.py                  # prove the workbook is still live
python scripts/test_rescale.py                     # prove it follows other numbers
python scripts/ibcs_doc.py --xlsx book.xlsx --check      # the guide agrees
```

The Excel build puts one template on each sheet of a single workbook, behind a
Read me. Pass a comma-separated list to `--template` for more than one, and
`--simple` for the base-tier version of each.

**`--doc` writes the by-hand build guide for that exact workbook** - how to
build it in Excel with nothing but the menus. Part 1 is the sheet skeleton every
template shares, Part 2 a worked build per family, Part 3 a per-sheet reference.

It is generated rather than written because every range, formula, hex colour,
gap width and axis bound in it is **read back out of the file just built** -
including how each individual *point* is painted, which is where the forecast
hatch and the pin head's scenario live. A hand-written guide drifts the first
time a formula changes and nothing catches it, because prose has no tests.
`--check` runs the other way, asserting that every formula an existing guide
quotes is still at the address it names.

**Four workbooks ship**, two datasets by two forms, all from the same code, each
with a generated guide beside it in `guides/`:

| File | What it holds |
|---|---|
| `IBCS Charts - Complex Versions.xlsx` | all 17 templates with every tier |
| `IBCS Charts - Simple Variants Only.xlsx` | the seven templates with a useful base-tier form - C01A, C02A, C03A, C04A, C05X, C06F, C12A - drawn with that tier only. A scattergram has no tiers to drop, a line chart's tiers are all measures, and a table reduced to one column block is a list rather than a report |
| `IBCS Charts-Progressive - Complex Versions.xlsx` | the same 17, on a real company's filed figures |
| `IBCS Charts-Progressive - Simple Variants Only.xlsx` | the same seven, likewise |

The Institute pair *is* the recreation and its cells hold the published figures.
**The Progressive pair is the one to paste over** - every figure in it belongs to
a real company, so none of them is a number a reader has to identify before
replacing. See **Drawing other data** below.

Each has a **Read me** sheet first, then one sheet per template in template-id
order, and **every sheet prints on one page**: print area set to the chart zone
alone, `FitToPagesWide/Tall = 1`, gridlines off. The data zone is deliberately
outside the print area - it is the input, not the deliverable.

## Drawing other data

A fidelity recreation makes a poor file to paste your own numbers into: you
cannot tell which figures are Furniture Inc.'s without checking each one. So the
seventeen are drawn a second time, from The Progressive Corporation's SEC
filings. `references/guide-second-dataset.md` is the whole pipeline; these are
the seams, and they hold for any dataset.

```
scripts/ibcs_data_alt.py    the seventeen Templates, read from datasets/pgr/
scripts/build_alt.py        the entry point - see below
scripts/fetch_pgr.py        harvests the filings from SEC EDGAR. Not a build step
scripts/pgr_parse.py        turns one release into figures, keyed on row labels
scripts/test_alt_ties.py    the tie-outs, and the gate build_alt passes to build()
scripts/test_pgr_reconcile.py  the harvested months against the audited filing
datasets/pgr/*.csv          five files; the figures, committed
```

```bash
python scripts/build_alt.py pgr.xlsx --doc "pgr - Excel Guide.md"
python scripts/build_alt.py out.xlsx --template C03A --export-dir pics/
python scripts/test_alt_ties.py            # 211 tie-outs
python scripts/test_pgr_reconcile.py       # 605 checks against the filed totals
python scripts/test_dataset_isolation.py   # no renderer reaches past its template
python scripts/test_provenance.py          # constructed figures are shaded (Excel)
```

**`build_alt.py` is separate from `ibcs_excel.py`'s CLI on purpose.** That CLI
resolves template ids against `ibcs_data`, so it cannot be pointed elsewhere.
`build()` has always taken the templates it is handed - nothing swaps a module,
nothing monkeypatches, and that is the property the next rule protects.

**Nothing is read from the data module by name.** Both renderers once reached
past the `Template` they were given and read module globals - C08H's opening
balance, C09C's accents and iso-curves, C13D's panel figures, plus C10D's
suppressed labels and C11A's pixel rulers on the SVG side. That fails *silently*:
a global left behind does not raise, it draws the first dataset's figures on a
sheet titled with the second one's entity. `test_dataset_isolation.py` asserts
the property against the live module rather than a list of forbidden names, so a
new one fails the day it is written.

**Say where every figure came from, and be finer than "real or invented".**
`Template.provenance` carries `SourceNote(scenario, basis, detail)` with three
bases - **filed**, it appears in a source document; **derived**, computed from
filed figures by a rule; **assumed**, a modelling choice nobody published. A
company's stated target is filed. Applying that target to a month to make a plan
figure is assumed: the same number, a different claim.
`references/reference-provenance.md`.

**Provenance is an Excel facility. The SVG renderer has none.** `ibcs_svg.py`
reads no `SourceNote`, applies no assumed fill and writes no Sources sheet.
Never promise a web render that distinguishes a constructed figure from a
reported one.

**Write real tie-outs, and gate the build on them.** `check_ties()` reads
`ibcs_data` module globals, so handed another dataset it re-verifies the
Institute's figures against themselves, passes, and proves nothing. `build()`
takes a `check_ties` parameter for this, and the gate passed in must **refuse a
template it has no checks for** - otherwise a sheet can be added without its
arithmetic ever being checked.

**Real data pays for the work in reconciliation.** Figures from independent
documents that must agree - twelve monthly releases summing to an audited
annual, two tables in one filing carrying the same number - catch transcription
errors that no self-consistency check can, because each one produces plausible,
wrong numbers rather than an error.

**A workbook of a real company's figures must say it is not that company's
reporting.** Polished notation on public figures reads as house reporting, so
generated output carries two disclaimers: the IBCS® one, which says nothing
about the company, and a non-affiliation statement. Anything built by hand from
such a dataset has to reproduce both.

## The template library

Seventeen templates: C01-C13 charts, T01-T04 tables. `references/decisions.md`
records which variant was chosen for each and why.

**The IBCS® reference renders are not distributed with this skill.** They are
the IBCS Institute's own images of all 63 published variants, and they are used
here only to measure a recreation against its source. `compare_render.py` and
`extract_palette.py` are the only two scripts that want them; point
`$IBCS_TEMPLATE_REFS` at your own copy, or skip those two - nothing else needs
them. The recreation is built from `ibcs_data.py` and the Progressive
workbooks from `ibcs_data_alt.py`, and neither renderer reads an image.

**All seventeen build end to end in both engines**, in SVG and in Excel:
**C03A** (multi-tier columns, dPY), **C04A** (multi-tier bars, dPL), **C05X**
(columns bridged to plan, with a forecast half), **C06F** (bars bridged to prior
year), **C12A** (P&L waterfalls), all four tables - **T01B** and **T02A** a
hierarchy printed and drawn, **T03A** and **T04A** a statement printed and drawn
- and the stacked structure pair, **C01A** (three panels on one scale) and
**C02A** (bars with the totals row as legend), plus **C07C** (cumulative lines
over a monthly tier) and **C08H** (a stock over its movements), and the XY
pair - **C10D** (a product-market portfolio in bubbles) and **C09C** (149
products against iso-gross-profit curves), **C11A** (an ROI tree: six small
charts and the arithmetic between them) and **C13D** (fifteen locations as
small multiples, built as one chart object by the `panel-charts` skill).
Fidelity is measured, not judged:
C12A's eighty bars land within 4px of the original, C06F's forty-five and C05X's
bridge within 3px, T01B's 280 cells and T03A's 168 within 3px, T02A's 77 drawn
elements within 3px, T04A's 38 within 4px, C01A's column tops within 1px and
C02A's bar ends within 4px, C07C's monthly tier within 3px and C08H's flow
columns within 2px, C10D's twenty bubbles within 3.3px in position and 2.2px in
radius, and C09C's three curves within 1px with 144 of its 149 markers exact to
the centre pixel, C11A's 48 drawn elements within 2px with 46 of them within
1px, and every one of C13D's 181 within 1px. `references/decisions.md` records
how each was measured and where each diverges.

**C13D is drawn two ways on purpose.** The SVG follows the reference - Berlin
spanning two rows, the location average as columns on a shaded ground - and the
workbook draws a uniform 4x4 of pins as a **single native chart object**, built
by the `panel-charts` skill. Excel has no panel chart type, and the workaround
of copying one chart sixteen times leaves sixteen charts whose axes drift apart
the first time a figure changes.

Between them they cover most of the primitives: the tier stack in both
orientations, scenario fills and the FC hatch, variance colour by impact, pins,
the PL/BU double rule, totals that share the category axis with the categories,
and both kinds of waterfall - the accumulation that builds a statement from zero
and the bridge that carries one total onto another, in both orientations and
with forecast steps - plus the table grid, charts aligned to its rows, the ratio
row that a statement reports but does not add up, the stacked structure chart in
both orientations, lines with scenario markers, an area over a stock, and XY
plotting in both its forms - a size channel with its own legend, and a
categorical colour with a key - the driver tree, where the connectors state a
calculation and boxes sharing a unit share a scale, and the small-multiple grid,
where fifteen panels share one ruler and one chart object. Nothing in the
template library is unbuilt.

## Reference map

| Read this | When you need |
|---|---|
| `references/decisions.md` | Which IBCS® variant was chosen for each template and why, how each recreation's fidelity was measured, and the Excel traps that cost the most time |
| `references/guide-second-dataset.md` | **Before drawing anyone's real figures.** The entry point, the harvest and its SEC user agent, the CSV contract, the four gates, and a checklist for a third dataset |
| `references/reference-provenance.md` | `SourceNote` and the filed/derived/assumed taxonomy, what the marking changes on a sheet, and the attribution a real company's figures oblige |

## Related

[`panel-charts`](https://github.com/wfphillips128/Panel-Charts-for-Excel-via-Claude-Skill)
builds a grid of small charts in Excel as one chart object. C13 is its first
caller and the two are designed to fit: that skill supplies the geometry, this
one supplies the notation. It has to be installed beside this skill for C13D to
build from source - see "The one dependency worth explaining" in the README.

`viz-design` covers general chart craft and non-IBCS® palettes. Use it when the
work does not need a formal notation standard; use this skill when it does.
