# Decisions and findings

Why this skill is built the way it is, and what Excel does that the documentation
does not mention. Recorded so the next seventeen templates do not re-derive it.

---

## Palette provenance

IBCS v2.0 rule UN 4.1 (p52): *"Although IBCS does not prescribe exact codes for
these colors, they should still be consistent within an organization."*

So there is no canonical palette to recover, and hex sets circulating online
(Zebra BI, the Plotly community thread) are other houses' choices - valid, but
not normative. Mining the standard's PDF figures was tried and abandoned: they
are raster and yield decorative noise.

What `assets/ibcs-palette.json` contains is sampled from the IBCS Institute's own
published template renderings, using `scripts/extract_palette.py`. That makes it
a house palette that happens to reproduce the originals pixel for pixel, which is
what a fidelity proof needs. Documented as a house choice, not a rule.

The FC hatch geometry was measured rather than guessed: dark runs of 3px on an
8px pitch along a horizontal scanline, with the phase shifting 1px left per row
down - 45° exactly, ascending left to right. Note that a 45° stripe crossed
horizontally reads √2 wider than it is, so the SVG pattern divides both pitch and
stroke by √2.

One inconsistency worth knowing: C03A renders AC as `#3C3C3C` while the other
four sampled references and both 2026-dated renders use `#404040`. The house
value is `#404040`; the difference is not perceptible.

**PL was wrong and is now white.** The palette carried `#F2F2F2` for the PL and
BU body, justified in its own note as "what the renderings actually use". It is
not: C04_04A's PL bars are `#FFFFFF` inside, and C06_06F and T04_T04A contain no
`#F2F2F2` pixel at all. The reasoning attached to it - that a true white body
reads as a hole where series overlap - was plausible and unmeasured, and C04A
disproves it directly, drawing PL over AC in white with the outline doing the
delimiting. FC keeps `#F2F2F2`, which is its hatch background and *was*
measured. The lesson is narrow and worth keeping: a note saying a value was
measured is not the same as the value having been measured.

---

## Transcription policy

The reference renders show rounded labels over decimal source data. You cannot
recover PY by subtracting the rounded absolute variance from the rounded actual -
two roundings compound, and in C03A the error reaches 3 kEUR, enough to draw a
visibly wrong grey bar.

Instead, transcribe the two independently readable series - the measure (integer
labels) and the relative variance (one decimal) - and derive the rest:

    PY = AC / (1 + ΔPY%)

This was checked against the drawing, not just against itself: predicted PY bar
tops match the reference's grey bars to within 1.2px across all nine actual
months. `check_ties()` then proves the derived absolute variance rounds back to
the labels IBCS printed, with a documented 1.0 tolerance because a derived 5.5
can legitimately print as 5 or 6.

**Print what the original printed.** Where a template records IBCS's own label
text, use it. The bar keeps its exact derived geometry; only the caption changes.
A recreation that argues with the original about its own numbers is not a
recreation.

Two templates now carry independent tie-outs. C03A's is internal - derived PY
reproduces every printed variance label. C04A's is stronger and genuinely
external: plan is derived row by row as `AC - ΔPL`, and the nineteen derived
values sum to exactly the 1 889 kUSD total IBCS prints separately and which is
never fed in. A transcription error in either input series would break that sum.

---

## The workbook must stay live

The shipped workbook is charts and data, and someone with only the file must be
able to retype a figure and get a correct chart. That constraint is stronger than
it first appears, because a chart built the obvious way looks identical on the
build data and starts lying the moment anyone edits it.

**Two typed columns, everything else a formula.** PY and the measure are inputs;
the variances and every chart-feed column derive from them. Full precision in the
cells, no pre-rounding - pre-rounding a variance to one decimal turned November's
−12.4926 into −12.5, which Excel then rounded away from zero to −13 where IBCS
printed −12. That single `round(..., 1)` was the whole cause of a label
discrepancy previously patched with static text.

**Colour has to come from the series, not the point.** Excel cannot colour a
point by its value, so a variance is plotted as two mutually exclusive series -
`=IF($E6>0,$E6,NA())` and its negative twin - each solid in one colour. Retype a
figure so the variance flips sign and the bar moves across the axis *and* changes
colour on recalculation. Per-point formatting cannot do this; it is frozen at
build time.

Use `NA()`, never `""` or `0`: a blank can coerce to zero and a zero draws a
flat bar sitting on the axis.

**A third, invisible series carries the labels.** Excel creates a data label only
where a value exists when labels are switched on. The two coloured series each
hold half the points, so a category that is empty at build time and gains a value
after an edit would draw a bar with no number against it. A hidden series holding
the *complete* variance has a value everywhere, so every label exists from the
start and simply follows the data.

**Clustered at Overlap 100, not stacked.** Both put the series in one slot, but a
stacked chart refuses every outside label position - Excel offers only Center,
InsideEnd and InsideBase - and IBCS puts value labels outside the element.

**Scale is fixed deliberately, and that is also what gives headroom.** Value-axis
bounds are chosen so tiers sharing a unit share a scale, which IBCS requires:
the measure tier sets the exchange rate at 230 kEUR over 210pt, so the
absolute-variance tier's 84pt must span 92 kEUR. Sizing that tier to fit its own
data instead would draw variances larger than the measures they came from, which
is exactly what the rule prevents. `verify()` now checks it, because the mistake
produces a chart that looks fine.

It has a second benefit: the correct range is far wider than the data needs, so
edits stay on the plot. The first version of this, sized to fit, clipped at −20
and a test edit to −33 ran off the chart taking its label with it - which looked
like a label bug and was a scaling bug.

**Known limit, documented rather than solved.** Scenario fill on the measure tier
stays per-point and static. Making AC and FC separate series would force a third
cluster slot and shift every bar. Values respond; which periods are forecast is
structure, not data.

`scripts/test_responsive.py` proves all of this end to end - it edits a cell,
recalculates, and samples the exported image for the colour actually rendered.
Every built template now carries a claim there, and each one is the claim that
template's chart actually makes:

| Template | What an edit must do |
|---|---|
| C03A | recalculate the variances and change the bar's colour with the sign |
| C03A | a retyped unit reaches the subject line and every tier caption |
| C05X | a forecast month moves the plan, the total and the expected half - and not the measured half - while the bridge keeps landing on the closing column |
| C06F | a state moves every level below it and the bridge still lands on the actual total |
| C12A | a line re-floats every bar below it and its subtotal follows |

The C05X case is the densest, because one month feeds four places at once: its
own column, the plan column it is measured against, the closing column's expected
half, and the bridge that has to keep landing on that column. A workbook where
any one of those lags is worse than a static picture, because it looks current.
Sample the image using the locked plot geometry, not a fraction of its width: the
plot area is inset and offset, so a naive fraction lands on the neighbouring bar
and reports a false failure.

### And so must the words

The same requirement, applied to the text: nothing a chart says may be typed
into the chart. A workbook whose numbers respond to an edit but whose subject
line still reads "in kEUR" after the sheet was redenominated is not live, it is
half live, and the half that is wrong is the half that states the unit.

**The unit is its own cell, not a fragment of a sentence.** The first version
wrote row 2 as `f"{measure} in {unit}"` - one string built in Python. That is
editable but not *composable*: nothing could refer to the unit alone, so a tier
caption wanting "dPY kEUR" had to repeat it as a literal, and the sheet would
have had four places to change and no source of truth. The block is now six
typed inputs - entity, measure, unit, period, reference, message - and the lines
the reader sees are formulas over them.

The reference scenario is an input for the same reason. It reads as structure
rather than data, but it is in the text of every variance caption, so PL to BU
should not mean editing three text boxes.

**Two hops, one source.** The caption cell holds a formula (`="D"&$B$5&" "&$B$3`);
the text box on the chart links to the cell. Text boxes cannot hold formulas of
their own, so the cell is what makes the composition possible and the link is
what puts it on the chart.

**Do not set NumberFormat "@" on a cell you are about to give a formula.** It is
right for the typed inputs - a period of "2025" is a label, and left alone Excel
stores it as the number and right-aligns it - and it is silently wrong for the
derived ones, where it makes Excel display the formula instead of evaluating it.
That would put `=B2&" in "&B3` on the chart, looking like a broken link rather
than a formatting mistake.

**Where a caption goes is an orientation question**, and both answers are margin
the plot already reserves: a column stack labels its tiers in the left inset,
level with the zero line, which is where the SVG puts them; a bar stack has no
left margin to spare and uses the headroom above the plot. Both are positioned
from the *locked* plot geometry, for the same reason the double rule is - a
caption naming the reference scenario has to sit on that tier's zero line, not
near it.

**Both halves are asserted, and both assertions were made to fail before being
believed.** `check_captions` compares the caption the cells will produce against
the label the SVG prints, so the two renderers cannot drift apart while each
looks right alone. `check_caption_placement` measures where the box landed
against its tier's chart object and plot area; pushing the C03A captions 60pt
right made it report both tiers and exit non-zero. And `test_responsive.py`
retypes the unit cell and reads the text back off the shapes: the subject line
and the dPY caption follow, and the dPY% caption deliberately does not, because
a percentage carries no unit.

The row plan is computed from the block rather than declared. It used to be
`first_row = 6` with a comment saying rows 1-4 were the title, which stopped
being true the moment the block grew - and a stale constant there puts the data
under the headers silently.

---

## SVG must survive being embedded

The website project may place several charts on one page, so each file has to be
inert in company.

**Namespace every generated id.** SVG ids are global to the document. Two charts
that each define `hatch-404040-F2F2F2` collide, and the later one silently
repaints the earlier one's forecast bars - a failure that looks like a rendering
quirk. Ids are prefixed with the template id, and
`compare_render.py --grid` writes an inline page that would expose a clash.

**Keep width, height and viewBox.** The explicit size gives a bare `<img>` an
intrinsic dimension; the viewBox lets it scale into whatever cell a grid gives
it. Dropping either breaks one of the two cases.

---

## Excel findings

Everything below cost real time to find. Each fails with a bare
`Exception occurred` naming nothing.

**GapWidth caps at 500.** This kills the obvious pin implementation. On a
twelve-category axis, 500 bottoms out around a 9px stem where the reference uses
5px. Values above 500 raise.

**Pins are error bars.** The working technique: a line-markers series whose
square marker is the pin head, with a custom Y error bar as the stem. Error bar
line weight is set in points, so the stem width is exact. Two series are needed -
desirable and undesirable - because an error bar takes one colour per series.

**`Series.ErrorBar` must be called positionally.** pywin32 does not resolve its
named arguments; `ErrorBar(Direction=..., Include=...)` raises.

**`ErrorBars.EndStyle` takes `xlNoCap = 2`**, not `xlNone`. Passing `-4142`
poisons the chart, and every subsequent `PlotArea` call then fails too - which
sends you hunting in entirely the wrong place.

**Plot-area assignment order matters.** Excel validates each assignment against
the current state of the others, so setting a left edge while the auto-sized
width is still wide asks for a plot area running off the chart. Assign width
before left, height before top.

**The first plot-area lock can fail and the identical call then succeed.** Excel
defers chart layout, so an assignment made immediately after the series are
formatted may be rejected against geometry that is about to change.
`lock_plot_area` retries four times, re-fetching `PlotArea` each pass rather than
reusing a cached reference, and only raises on the last.

**Do not derive plot height from `ChartArea.Height`.** Excel quietly resizes a
chart object after creation - ask for 110pt, get 104pt - so anything computed
from it drifts between tiers. Pass the height in.

**On a line-type series, `Points(i).Format.Line` is the connecting segment, not
the marker border.** Styling markers through it makes stray lines appear between
adjacent pins. The marker border is `MarkerForegroundColor`.

**Excel and Python round halves differently.** Excel takes -12.5 to -13, Python
rounds to even and gives -12. So when overriding a label, compare against the
text Excel has actually rendered, not against what Python would render -
otherwise you miss exactly the cases the override exists for.

**A chart object moves with the cells under it unless told not to.** The default
`Placement` is move-and-size-with-cells, so widening a column drags the charts
sideways. Set `Placement = xlFreeFloating` (3) or any later column change
silently undoes a layout.

**A pattern fill's phase is anchored to the shape's position on the sheet.**
Move a chart and its FC hatch stripes land a pixel or two differently, because
the tiling is computed in sheet space rather than shape space - and since the
stripes run at 45 degrees, moving a chart *vertically* shifts them horizontally
too. Nothing about the fill has changed; only where it starts. This matters for
regression testing: a pixel comparison of exported charts is only meaningful
between builds that place the chart identically. Proved rather than assumed -
rebuilding at the old left *and* top reproduced the pre-move exports exactly.

**A text box linked to a cell needs the sheet name quoted.**
`DrawingObject.Formula = "='C03A'!$B$1"` works; `"=C03A!$B$1"` raises *"Unable to
set the Formula property of the TextBox class"*, which reads like the property is
unavailable rather than like a complaint about syntax. Worth knowing, because
the linked form is what lets a title block have one source of truth.

**Merged cells do not auto-fit their row height.** Wrap text in a merged range
and Excel will happily hide the overflow, so any merged block that can wrap needs
its height set explicitly.

### The tier stack

Five templates need two to four charts sharing one category axis. Excel will not
do it: autolayout sizes each plot area to fit that chart's own axis labels, so a
tier whose labels are one character wider silently shifts its bars sideways.

The approach that works:

* identical `Left` and `Width` on every tier chart;
* the value axis suppressed entirely, so nothing competes for horizontal space;
* `PlotArea.InsideLeft` / `InsideWidth` set explicitly to the same numbers after
  the series exist, and set again after all tiers are built, because adding a
  later chart relayouts earlier ones;
* `verify()` measures what Excel actually did and fails if the tiers disagree by
  more than half a point.

**A clustered column series sits inside its category slot; a line/marker series
sits on the category midpoint.** The measure tier puts PY on the midpoint with
the measure to its right, so column tiers align by carrying an invisible spacer
series that holds PY's slot, with the same gap and overlap. The pin tier cannot
do that - markers ignore clustering - so its plot area carries a deliberate
offset of one category × 9/56. `verify()` knows the intended offset, so the shift
passes while an accidental drift of the same size would not.

Measured result: PY occupies 0.16 of the category pitch, against 9/56 = 0.161 in
the reference.

### Bars, and the double rule

C04A is C03A transposed: tiers side by side sharing a vertical category axis,
measured against plan rather than prior year. Porting it forced three things.

**The layout fields are named for the shared axis, not for x and y.** A tier's
`offset`/`extent` is Top/Height when tiers stack down the page and Left/Width
when they sit side by side, and `plot_extent` is the corresponding InsideHeight
or InsideWidth. Naming them `top` and `height` and then using them as Left and
Width in the horizontal case is the kind of thing that reads fine and costs the
next session an hour.

**The double rule cannot be an axis.** Excel gives an axis one line, and the
notation needs two - PL and BU are fictitious scenarios drawn outlined rather
than filled, and an outlined bar seen edge-on is a pair of lines. That is what
lets a reader tell a dPL tier from a dPY one without a legend. So the axis line
is switched off and `draw_double_rule` puts two line *shapes* where it would
have been, positioned from the locked plot geometry and the value bounds so they
land on the data zero rather than wherever Excel put its own axis. It has to run
after the plot areas are locked, because it reads them.

The previous version drew a single line and renamed the chart
`"... [needs double rule]"`. That is worse than it sounds: a single line is not
a weaker version of the notation, it is the notation for a *different* reference
scenario. Wrong, not incomplete.

**Bar charts plot the first category at the bottom.** Every bar and table
template reads top-down, so `ReversePlotOrder` is required, and the value axis
crossing has to move with it or the category labels jump to the far side.

**Known limit: the pin head is not drawn on a bar chart.** The vertical pin is a
line-markers series with a custom Y error bar as the stem and the series' own
marker as the head. Excel refuses to combine a bar series with a line series, so
on a horizontal tier there is nothing to hang the error bar on and no marker to
use. The stem is a very thin bar instead - which is what makes the GapWidth
ceiling of 500 matter - and the head is absent. Drawing nineteen squares as
shapes was rejected: they would be static, and a head that stops tracking its
pin after an edit states something false, which is worse than one that is
missing. The fix, when it is worth it, is a stacked bar - an invisible segment
of `|value| - head` then a visible one of `head`, which puts a live square at
the tip carrying AC's fill.

### Waterfalls

C12A is a profit and loss statement drawn as two waterfalls with the variance
between them. It is the first template whose categories are not
interchangeable, and almost everything below follows from that.

**A statement's row carries three facts, and one number supplies all of them.**
Whether a line adds or subtracts decides where its bar starts, which of the two
shades it takes, and which direction counts as good. Holding those as three
lists would let them disagree; holding one `sign` per row means a cost line is
drawn light, floats downward and turns red when it rises, and none of those can
be true without the others. `Row.higher_is_better` is literally `sign > 0`.

**The walk lives in the data layer.** `waterfall_spans()` returns where each bar
starts and ends, and both renderers call it, because the rule has three cases
and only the first is obvious:

* an element moves the running total and spans the move;
* a subtotal with `spans="zero"` runs from the origin to the total - a *level*;
* a subtotal with `spans=n` runs across the n rows above it - a *sum*.

Neither kind of subtotal advances the total. One that did would count its own
components twice, and the error is invisible on every row above the first
subtotal - which is where you would look.

**Impact colour needs four series, not two.** The two-series split that makes a
variance colour follow its sign cannot express a statement's dPY column, where
favourable and adverse bars sit on the *same* side of the axis - revenue up and
cost up are both positive, and one is green. Excel colours a series, and the
value decides the side, so the four combinations of (side, impact) need four
series. `IMPACT_SPLIT_FORMULAS` reads the sign column, so re-classifying a line
from a cost to a credit recolours it without moving it. Per-point formatting
cannot do either.

**Subtotals must be formulas, and the failure is invisible.** They were typed
inputs at first, transcribed like everything else, and every check passed: the
bars were right, because a bar is drawn from the walk. What was wrong was the
*number printed beside it*. Retyping a component left its total unchanged, so
the sheet showed five expense lines summing to a total that no longer matched
them, and the dPY on that total was stale too. `typed_override()` now derives
every subtotal from the level column. The general lesson matches the zone rule's:
the geometry check could not see this, because the geometry was never wrong.

**No connectors in Excel, and the obvious answer is a trap.** Excel joins a
stacked chart's segments across categories - `ChartGroup.HasSeriesLines` - and
that is the documented way to build a waterfall. On a *floating* one it joins
the wrong pair of points: a series line runs from a segment boundary in one
category to the same boundary in the next, and once the base moves, that
boundary is at a different value in each, so every connector comes out diagonal.
The rendered proof is worth keeping in mind before reaching for it again.

A vertical connector cannot be made from stacked series at all. It needs a
boundary equal to the running level in both the row above and the row below, and
each row supplies one value per series. Drawing them as line shapes was rejected
on the rule this project has now applied three times: a shape does not follow the
data, and a connector left at the level the statement used to reach states
something false. The SVG carries them; the workbook shows the floating bars.

**Known limit: a stacked segment's label sits inside it.** Excel offers a stacked
series only Center, InsideEnd and InsideBase - it rejects OutsideEnd, which is
where IBCS puts a value - so a segment narrower than its own label clips it, and
C12A's 6, 45 and 43 print as slivers. The label text itself is live and correct:
a waterfall segment's value *is* the magnitude the reference prints, so nothing
is written onto a label here. Writing text onto one is what freezes it, and the
prior project's builder does exactly that. The fix, when it is worth a session,
is a thin invisible spacer series after the bar with its labels bound to the
value column through Value From Cells, which is a range reference and stays live.

Label colour is chosen from the luminance of the fill underneath rather than
from a list of dark colours - the palette is a house choice and can be changed,
and a hard-coded list would go stale the moment it was.

**Formats before values.** A statement labels its result lines "= Sales revenue",
and a cell given that string before it has been told it holds text is parsed as a
formula. It showed `#NAME?` in the cell *and on the chart*, because the category
axis reads the same cell. Number formats are now applied to the block before
anything is written into it.

### What the waterfall greys mean

A waterfall needs two shades per scenario, which is one more than the scenario
notation supplies: within a panel, a line that adds and a line that subtracts
have to be distinguishable, and both are the same scenario. So each panel takes
a two-step slice of the structure ramp and the panel header carries the
scenario - AC at `#404040`/`#7F7F7F`, PY at `#7F7F7F`/`#BFBFBF`, all measured
off the reference.

That makes C12A's PY panel one ramp step darker than the house PY of `#A6A6A6`.
It is not a contradiction: C06_06F and T04_T04A, which have no deduction shade to
accommodate, both use `#A6A6A6`. Read it as what a waterfall needs rather than as
a second opinion about PY.

### Outlier markers

Rule UN 5.3, confirmed in the standard rather than inferred: where a relative
variance is enormous because its base is tiny, mark it instead of rescaling every
other row to it - *"omit the pin head and add outlier triangles pointing in the
direction of growth"*. The head is omitted because a head states a measured value
at a position, and that position is not where the value is. Rule CH 4.4 gives the
reason: an outlier that is not important to the business should not distort the
scale of the chart.

The *number* of triangles is not in the standard. C12A draws three for +983% and
one for +391%, which is one per panel-width of overrun, and that is what the SVG
reproduces - an observed house choice, recorded as one.

In Excel the value axis bound does the clipping on its own, so an outlier bar
stops at the panel edge and stays live; the triangles are absent for the same
reason the connectors are.

### C05X: the densest notation in the library

C05X is the third waterfall and the one that earns the tranche. Its bridge runs
horizontally across twelve months from the plan column to the actual-plus-
forecast column, and its last four steps are *forecast* - so a step is coloured
by impact and filled by scenario at the same time. An expected improvement is
hatched green; a measured one is solid green. Nothing else in the library asks
the two systems to operate on one element.

**Its data is the cleanest in the project.** Every identity closes exactly, with
no rounding interval anywhere: the twelve monthly figures sum to the closing
column, split correctly into the printed measured and expected blocks, their
variances sum to the printed total variance, the derived plan sums to the printed
plan column, and all twelve percentages match the printed labels to one decimal.
Eight tie-outs, none of them needing a tolerance.

**One scale below the pin tier.** The two opening columns, the twelve monthly
columns, the bridge and the closing column are all 1.5975 px per kEUR - measured
off the two totals themselves - which is what lets a month's step be read against
the month's own column directly beneath it.

**The two closing variance bars are drawn to different widths, and it is not
decoration.** The wide one lies along the prior-year guide it is measured from;
the narrow one drops down the plan guide. Each reads against the line it belongs
to, which is why the template can show two comparisons of the same column
without labelling either.

**A third disagreement about greys.** C05X draws its prior-year column at
`#808080`, its monthly actuals at `#000000` and its plan fill at `#F2F2F2` -
where the house palette has `#A6A6A6`, `#404040` and `#FFFFFF`. That is now three
renders (with C03A's `#3C3C3C` and C12A's `#7F7F7F`) disagreeing with each other
about the same semantic. The house palette is used and the divergence recorded;
the palette's claim to reproduce the originals pixel for pixel holds for the
templates it was sampled from, not for every render IBCS has published.

Worth noting for the palette: C05X's plan fill is `#F2F2F2`, which is the value
that was removed from the palette after C04A showed white. Both occur. What is
binding is that plan is *outlined*, and the house choice of a white body stands.

### A waterfall that bridges rather than accumulates

C06F is the second waterfall and it walks differently. C12A's starts at zero and
builds a statement; this one starts on the prior-year total and lands on the
actual total, one step per state, with both totals drawn as bars beside it. So
`waterfall_spans` takes a starting level, and what it walks is the *variance*,
signed, rather than a magnitude with its direction held separately.

That is the distinction to carry into C05X and T04: a waterfall is either an
accumulation or a bridge, and only the accumulation needs the sign column.

**The transcription direction matters, and the wrong one is plausible.** C06F
prints AC, dPY and dPY% - it is over-determined, so a direction has to be
chosen. Deriving PY from AC and the printed dPY reproduces the reference's own
dPY% labels badly: Missouri comes out at -14 where IBCS prints -16, because a
small state's percentage is very sensitive and dPY is printed to whole units.
Deriving the other way - `PY = AC / (1 + dPY%)`, as C03A does - reproduces all
fifteen printed dPY labels to within 1. The rule generalises: derive *from* the
series whose printing loses the least information, which is rarely the one in
the biggest type.

**The tie-out is the bridge itself.** Fifteen derived variances have to carry the
printed PY total onto the printed AC total, and neither total is an input to the
derivation. They land on 1 727.3 against a printed 1 728.

**The shortfall bar.** Each state draws AC solid and then appends `PY - AC` in
PY's own light grey where prior year was higher, so the bar reaches PY and the
grey is exactly what was lost. It is what makes Illinois - which lost two thirds
of its sales - readable at a glance, and it is why the message is about Illinois.

Two things about it were got wrong first and are worth keeping:

* **Draw all the bars, then all the labels.** Drawing each row's label before the
  next element meant the shortfall was painted over the number it belonged to.
  It showed only on rows that had a shortfall, which made it look like a font
  problem rather than an ordering one.
* **The actual's number sits against its own bar**, even where the grey runs on
  past it. The reference prints California's 257 over the grey rather than clear
  of it: the number belongs to the dark bar, so it goes where the dark bar ends.
  Only the prior year's own label is conditional on having room, which in this
  variant means Illinois and nothing else - decided by measuring the segment
  rather than by naming the state.

**A sub-pixel step still has to be visible.** Four states moved by 1 to 3 kUSD,
which is under a pixel at this scale, and a step that rounds away leaves a gap in
the bridge - the reader sees the walk stop. Three pixels minimum, not one,
because a sub-pixel rectangle renders as two half-covered pixels and reads as a
smudge. The reference draws three.

**The reference disagrees with itself once.** Illinois' prior year derives to 457,
and the reference prints 456 - but its own printed 169 and -288 also give 457.
Ours is kept and the difference recorded, which is the same treatment C03A's
"prevoious" typo got: a recreation follows the original's numbers where they are
its own, and does not adopt an arithmetic slip.

**Re-render before measuring.** `compare_render.py` reads the SVG from disk; it
does not call the renderer. Two rounds of measurement were spent on a stale file
before that registered, both times producing differences that looked like real
geometry errors. Run `ibcs_svg.py --template X` first, always.

### The zone rule

A sheet has two reserved zones: a data zone of declared width on the left, and a
chart zone to its right. Nothing straddles them, and the chart zone's left edge
is *measured* from where the data zone ends rather than declared.

This exists because the first C03A workbook got it wrong in a way that was
invisible to every check it had. The builder called `Columns("A:N").AutoFit()`
with the IBCS message sentence sitting in F1, so AutoFit sized column F to hold
140 characters - 544pt, against 250pt for A to E combined - and the charts,
pinned at a hard-coded 300pt, landed inside it. The message, the whole dPY%
column feeding the top tier, and every chart-feed column ended up underneath the
charts or pushed past them.

Each half of that was defensible alone. Together they meant the layout depended
on the length of a prose string, which is not a layout at all.

What the fix actually consists of:

* **column widths are declared, never AutoFit** - one long cell must not be able
  to move the boundary;
* **the chart zone starts where the data zone stops**, from `Range.Width`;
* **`Placement = xlFreeFloating`** on every chart and text box, or a later column
  change re-introduces the overlap;
* **the title block is merged** across the data zone. Text in an unmerged cell
  spills across its empty neighbours, and here the spill would run under the
  charts - invisible, because Excel draws shapes over cell text;
* **`verify()` asserts it**, comparing each shape's `TopLeftCell.Column` against
  the first column past the data zone. This was checked by forcing the charts
  back to 300pt: all seven shapes were reported and the build failed.

The last point is the general lesson. The charts had always been *correct*; only
where they sat was wrong, and nothing in a chart-geometry check can see that. An
assertion that only measures the thing you were thinking about will not catch the
thing you were not.

The title block is a live link, not a copy: the four text boxes over the chart
carry `='C03A'!$B$n`, so the cells in the data zone remain the single source of
truth and editing one updates the chart. The cells are labelled Entity, Subject,
Period and Message in column A - without those labels the same four lines appear
twice on screen and read as a duplication rather than as source and render.

---

## Template variant picks

**This table is authoritative.** `template-spec.md` in the project directory
carries an earlier version of it whose C07, C08, C09, C10 and C13 rows were
picked from page text before the variants had been looked at; all five changed
once they had been. Where the two disagree, this one is right.

**Complex** = a faithful recreation of a named IBCS variant, all tiers and
annotations. **Simple** = the same chart and data reduced to its base tier: no
variance tiers, no callouts, no highlight markers. IBCS's own variants do not
form a complexity gradient - they vary by scenario pair, sort order, measure
basis and precision - so simple/complex is a dimension we define.

| Template | Pick | Why |
|---|---|---|
| C01 Stacked columns | 01A | Title block, two structure panels, bracket callouts, comments |
| C02 Stacked bars | 02A | Europe/RoW grouping, subtotals, integrated legends, outlier arrows |
| C03 Multi-tier columns | 03A | 3 tiers, FC hatch, full-year block - the primitive-densest chart |
| C04 Multi-tier bars | 04A | Sorted by ΔPL; variants differ by sort only |
| C05 Columns + horiz. waterfall | 05X | ΔPL pairing, AC+FC stack |
| C06 Bars + vert. waterfall | 06F | Sorted by ΔPY, the IBCS-preferred order |
| C07 Line charts | 07C | 07A is ~750 unlabelled daily points; 07C is fully labelled |
| C08 Area charts | 08H | Quarterly inventory - all three scenario fills in one chart |
| C09 Scattergrams | 09C | Iso-gross-profit hyperbolae |
| C10 Bubble charts | 10D | AC vs PY paired bubbles; 14 labelled, transcribes exactly |
| C11 Tree charts | 11A | Only variant |
| C12 P&L waterfalls | 12A | Two waterfalls, ΔPY columns, ΔPY% pins with overflow arrows |
| C13 Small multiples | 13D | 16 panels of pins - exercises the pin primitive sixteen times |
| T01 Hierarchical table | T01B | Month and YTD blocks, subtotals, threshold footnote |
| T02 + integrated bars | T02A | Dual-panel month + YTD |
| T03 Measure rows | T03A | Full P&L, both variance pairs |
| T04 + integrated waterfalls | T04A | Waterfall column plus pin column |

**Carry forward:** C13 is ~190 transcribed values, the largest single
transcription job.

~~C09 has no labelled variant, so its ~130 points must be synthesized.~~
**Wrong, and corrected in the XY section below.** None of C09's variants prints
its values, but all 149 markers are *drawn*, and a drawn position is a
measurement. Every point in C09C came off the reference render; nothing in it is
invented, and no disclaimer is needed on the Read me sheet.

### What is actually built

Reference images exist for all 63 variants. Working renderers do not.

**All seventeen build end to end in both engines.** Built: C03A, C04A, C05X,
C06F, C12A, T01B, T02A, T03A, T04A, C01A, C02A, C07C, C08H, C09C, C10D, C11A,
C13D. The library is complete.

**C13D is built, and the separate skill it waited on is what built it.**
C13 *is* a panel chart, and Excel has no panel chart type - the gap was wide
enough that the general capability was built first, as the standalone
`panel-charts` skill, rather than invented once inside C13 and then again the
next time anyone needed a grid of charts - and C13 became its first
caller. The starting point was a set of the user's own notes on building a
panel chart in Excel by hand, with three worked workbooks beside them. When that skill exists, C13 becomes a
caller: it supplies IBCS notation and sixteen tiers of pins, and the panel
engine supplies the grid.

**C11A is built** - see the Tree charts section at the end of this file. It
needed a sixth sheet family: a tree is neither a tier stack, a table, a
structure panel, a line nor an XY plot, so `TreeLayout` / `TREE_LAYOUTS` join
the registries `check_registries` keeps in step.

Fidelity, measured: T01B's 280 cells land within 3px of the original with none
outside tolerance, its 15 columns within 1px and its 18 rules within 1.5px;
T02A's 77 drawn elements are all within 3px; T03A's 168 cells are within 3px
except the italic ratio row and two bold labels, where the rasteriser's
synthesised italic and bold are a few px wider than the reference's real ones;
T04A's 38 drawn elements are within 4px; C01A's seven column tops within 1px;
and C02A's nineteen bar ends within 4px with every Direct/Retail boundary
exact. Fidelity is measured: C06F's forty-five drawn elements and
C05X's twelve bridge steps land within 3px of their originals, C05X's two closing
variance bars are exact, and C12A's eighty bars land within 4px.

**C05X's measure tier carries one chosen divergence.** The monthly plan and
actual are clustered with an overlap so the plan sits behind the actual, and the
original stacks the closing column - measured under expected. One Excel chart
group is one or the other, so the workbook draws the closing pair side by side
and the SVG keeps the stack. The alternative was to stack the monthly pair too,
which would read as plan *plus* actual: a fidelity loss beats wrong notation.

### Porting the bridge to Excel

**Five fills on one axis means five series.** C06F's left panel draws a state's
actual, its shortfall in prior year's grey, and three scenario totals - and
Excel colours a series, never a point. So the panel is a stack of five, each
holding zero where it does not apply. Zero rather than NA(): a stack needs a
number in every slot, and NA() is for a bar that must not exist rather than one
with no length.

**A stacked segment has a length, not a sign, and the sign comes back through
the number format.** The bridge's two halves hold magnitudes, so a step of 288
would print as 288. Giving the adverse series the format `"-"#,##0` prints it as
-288 without anything being written onto the label - which matters, because
writing text onto a label is what freezes it. The same trick will serve any
stacked variance.

**A total that has components is a formula.** This is C12A's subtotal lesson
again, and this time the failure was visible: while the actual total was typed,
raising a state left the total where it was and the bridge stopped landing on
it - the chart said the walk ended somewhere the bar did not. `test_responsive`
now edits Illinois and asserts the closing level still equals the actual total,
which it did not before this was fixed.

The cost is two labels. The derived totals come to 1 730 and 2 074 where IBCS
printed 1 728 and 2 071, because fifteen rounded states do not add to a rounded
total. That is the split this project has settled into: **the SVG prints what the
original printed, the workbook prints what its own data says.** A workbook that
disagrees with itself is worse than one that disagrees with the reference by two,
and the same reasoning decides the variance labels - the reference prints New
York at -30 and the live sheet derives -29.

**Two bugs worth naming, because both drew something plausible.**

* Taking a variance bar's sign from its `span` rather than its `reference` gave
  the total dPY as +343. The bar was in the right place and coloured green,
  which is exactly the kind of wrong that survives a visual check.
* A variance-only row given a relative variance drew a pin as well as its bar,
  printing -17 twice - once against the actual total and once against the bar
  that states the same thing. A row that *is* a variance does not also get one.

**No vertical guides in Excel.** The reference runs lines down from the plan,
actual and prior-year totals so the bridge can be read against them. They would
have to be drawn shapes, and a shape sits where the total used to be after an
edit - the same reason C12A has no connectors. The SVG carries them.

### Porting C05X, and two dispatch bugs worth naming

Both bugs came from deciding *which builder to use* by looking at the data rather
than at the drawing, and both produced a bare COM error naming nothing.

**Stacked is not one constant.** `XL_BAR_STACKED` was written when both existing
waterfalls were horizontal, with no orientation test - so C05X, the first
vertical one, came out drawn on its side with the months running down the page.
Stacked, like clustered, has a bar form and a column form.

**Two panels can feed from the same column and want different charts.** The
choice of measure builder keyed on `"bar_py" in layout.col`, which is true of
both C06F and C05X - so C06F's *stacked* panel went into the clustered builder,
which asked Excel for an outside label on a stack and got `Exception occurred`
pointing at nothing. How a tier is drawn is what separates them, so `spec.stacked`
is what decides, and it is tested first.

The general lesson: dispatch on the property that actually differs. A column plan
is evidence about a template, not a specification of how to draw it.

**Two more places assumed a name would match.** A variance bar's `span` names two
summary rows and its `reference` names a scenario, and matching them by string
worked only while the rows happened to be called PL and PY - C05X's are "2025 PL"
and "2024 AC". `variance_ends` now asks each end what scenario it holds. And
`_typed_value` assumed a typed column and a tier share a name, which C05X breaks
by typing a `var_abs` column whose tier is the waterfall; `typed_source` declares
the mapping where it differs.

### Rows that aggregate the categories

C06F's sheet carries five rows that are not states: the plan, prior-year and
actual totals the states sit between, and the two variance bars beneath them.
They share the category axis in the drawing, so they have to share it in the
data, and `write_data` looped over `template.categories` and nothing else.

The answer was to generalise something that already existed rather than add a
parallel concept. Every template with a total already had a `Summary` - C03A's
separated full-year block, C04A's USA row - and they are the same object; only
the *number* of them varies. So `Summary` became `Template.summary_rows`, a
sequence, with `Template.summary` kept as the first of them so the two working
templates were untouched. Two fields were added:

* `span`, naming the two totals a variance-only row runs between. C06F's total
  dPY is a bar from the AC total to the PY total with no bar of its own, and
  naming the ends means the geometry follows the totals instead of repeating
  them - a bar cannot end up drawn against a total it is not a variance from.
* `before`, because C06F's plan and prior-year totals *lead* the states while
  C03A's and C04A's follow them. `Template.sheet_rows()` is the one place that
  decides the order, and it gives C06F `PL | PY | states... | AC | dPY | vs PL`,
  which is how the reference reads.

The C06F renderer was then repointed from the module constants it had been using
onto `summary_row(...)`, and the SVG came out **byte-identical** - which is the
proof a refactor wants: the same drawing, from one source instead of two.

The rule this establishes for C05X and the four tables: a total is data, it
lives in the data layer, and both engines read the same one.

C12A's fidelity is measured rather than judged: across all four panels, every
one of its eighty bars lands within 4px of the IBCS original at 1280px wide, and
all but one within 3px - the exception being the +983% outlier row, whose bar is
clipped to the panel edge in both drawings and so is not really a measurement of
anything.
Its transcription is the best checked in the project - a statement is a set of
arithmetic identities, so the six subtotals in each of two scenarios give twelve
independent tie-outs that no single mistyped line can survive.

Both earlier templates are faithful recreations of their reference renders. C04A's
SVG carries the parts Excel does not: the USA totals row, the reading rules every
five rows, the highlight ovals and the arrow, and the pin head markers. Those are
the next things to port if C04A's workbook is meant to stand alone rather than
sit beside the SVG.

What the three of them now provide, for the fourteen still to come:
the tier stack, the zone rule, scenario fills and the FC hatch, variance colour
by impact, pins in both orientations, the horizontal/bar code path, the PL/BU
double rule, and a live workbook whose typed columns differ per template. C12A
adds the waterfall itself - the walk, the two-shade panel, per-row impact colour
and derived subtotals - which C05, C06 and T04 all need, and a text block whose
every word comes from a cell.

---

## Tables

A table is a different code path from every chart, and the useful discovery is
how little of it is new. A column of a table turns out to *be* a `Tier` - a keyed
set of values optionally measured against a named scenario - so the whole data
layer carried over untouched, and a row hierarchy with subtotals is what
`Row`/`waterfall_spans` already described for a P&L. What had to be written was
the layout: a column model, a row model, and the header rules.

### The cumulation mark is notation, not a label

The second block of T01B and T02A is headed **`_November`**, not "YTD". UN 4.2:
a leading underscore marks a year-to-date period, a *trailing* one marks
year-to-go ("Jun 2026_"), and a leading tilde marks a moving annual analysis -
which is what C07 is labelling when it prints MAT.

Worth stating because it is invisible until measured. The block header's ink sits
at 929-1000 where "November" alone is 72px wide and the block centre is 960: the
string is "_November" centred as a whole, with the underscore hanging left of the
visible word. Reading it as a stray rule would have lost a rule of the standard.

### What a table's tie-outs are

The strongest of the project so far, because a hierarchical table states every
number it uses. T01B's 120 transcribed figures are covered by six subtotal walks
- three scenarios across two blocks - and a single mistyped country breaks its
own group and the World row and nothing else, which is what makes the failure
legible. On top of that, all 80 printed percentages are reproduced from the
derived values.

**The absolute variance columns are deliberately not transcribed.** AC - PY over
integers is exact, so a printed absolute variance would test nothing that its own
percentage does not test better. Transcribing targets that cannot fail is how a
tie-out suite gets long without getting stronger.

### The threshold is a cell, not a caption

T01B prints "Δ in red if: <-20, <-10%" and colours the cells that cross it. IBCS
requires the threshold to be stated wherever it is applied, and the only way to
make that statement *true* rather than decorative is to have one number:
`Tier.red_below` in the data layer, and on the sheet a footnote cell that the
conditional format itself points at. `test_responsive` widens that cell to -40
and asserts the red goes away; if it did not, the footnote and the colouring
would be two numbers that merely happen to agree today.

The tie-out for it derives the red cells from the stated rule and compares them
against the 23 the reference actually prints - keyed by row *index*, because
three rows are called "Other" and a label-keyed set collapses them into one. That
version passed while reddening the wrong continent's Other.

### A zero variance carries no sign

IBCS prints World's cumulative ΔPY% of +0.0215 as "0.0%". Excel says this with
the third section of a number format; Python format specs have no equivalent, so
`Tier.label_for` applies it. On the Excel side the plain third section is not
enough - it fires only on an *exact* zero - so the panels use conditional
sections: `[>0.05]+0.0"%";[<-0.05]-0.0"%";0.0"%"`.

Two other Excel format traps, both found by reading the cells back rather than by
looking at the sheet:

* `# ##0` gives the space separator IBCS prints, but emits the space
  unconditionally, so 59 comes out " 59" and a signed column reads "+ 59". The
  fix is conditional: `[>999]# ##0;[<-999]-# ##0;0`.
* `"<"# ##0` on -20 renders "-< 20". The minus has to be inside the literal:
  `"<"0;"<-"0`.

### Two sources that disagree

T01B and T02A are the same report told two ways, and their November blocks are
identical value for value - 40 figures transcribed once. **Their cumulative
blocks are not.** T02A's Other in Europe reads 5 899 against T01B's 5 441, and
T01B's China and Japan plans read 1 925 and 87 against T02A's 2 311 and 139; each
carries into its own subtotals and the World row, seven rows in all.

Both are transcribed as printed rather than reconciled, and `check_ties` asserts
the divergence is exactly those seven rows - so a later tidy-up that "fixes" one
to match the other fails loudly instead of quietly destroying a transcription.

### Drawn columns: what T02 and T04 add

A panel column is a `TableColumn` whose cells are drawn. `Template.panel_tiers`
names them, because that is precisely what separates T01 from T02 - the same
figures, stated as bars - so it is a fact about the template and both renderers
read it rather than each being told.

**The panels share a scale and differ in width.** Both of T02A's ΔPL panels carry
0.4472 px per kEUR and both ΔPL% panels 3.60 px per point; the cumulative ΔPL
panel is twice as wide because its numbers are twice as big. That is the right
way round and it is measurable: fitting each panel to its own range would draw a
month variance of 88 and a cumulative one of 211 at the same length. The
measurement lives in `ibcs_layout.PanelGeometry`, read by both engines, because
the zero rule's position is the one thing they must not disagree about - and the
Excel port needs neither number in pixels, only the zero *fraction* and the value
bounds, both of which fall out of it.

**The zero rule is the notation.** A ΔPL panel's zero is a double rule - plan is
drawn outlined, and an outlined bar seen edge-on is two lines - which is why
these panels need no legend to say what they are variances from. In Excel it is
drawn as shapes, and the chart's own category axis has to be switched off: a bar
chart draws that axis *at zero*, exactly where the double rule goes, so left on
it contributes a third line to a two-line notation.

### The pin's head belongs under its stem

The reference draws the stem *over* the head, so the pin reads its full length to
the value it marks. Ours drew the head last, covering the final few pixels of the
stem, and the pins stopped 4-5px short - 15 of 77 drawn elements out of
tolerance, all of them positive, which is the tell.

The fix is a reordering, and its safety is provable rather than argued: the set of
rectangles each template emits is unchanged, only their z-order. C03A and C05X
changed; C04A, C06F, C12A and T01B stayed byte-identical.

### Aligning a chart to table rows

The hardest thing in the table tranche, and it comes out exact for two reasons.

**Position from the cells, never from arithmetic over row heights.**
`Cells(...).Top` and `.Left` give the anchor; anything computed drifts.

**Ask, look, adjust.** A plot area the full size of its chart is not what you
get: Excel keeps a margin and returns 289pt inside a 300pt chart, which divides
twenty rows into bands of 14.45pt against a row of 15 - most of a row of drift by
the bottom, and invisible until someone reads a bar against the wrong country. So
the chart is enlarged by whatever the plot came up short and moved by whatever
inset Excel kept, and measured again; three passes settle it. `verify_panels`
asserts the band equals the row height and that the band *starts* where the rows
do, and it caught this on the first build.

The plot-area assignment order from `lock_plot_area` applies unchanged: size
before position, or Excel validates a left edge against an auto-sized width,
finds the plot hanging off the chart, and fails with a bare E_FAIL naming
nothing.

Two smaller ones. A panel is narrow and Excel wraps a data label to fit it, so
"+47" comes out as "+4" over "7" - there is no width to give it, because the
width *is* the scale, so the wrapping is what goes. And the chart must be
transparent: the plot is exactly its column's width but the chart around it is
wider, so an opaque one whites out the neighbouring column.

**Known limit, inherited from C04A:** a bar chart cannot carry the pin's head
marker. Excel refuses to combine a bar series with a line series, so there is
nothing to hang a marker on. The SVG draws it; the workbook does not.

### Statement tables: the ratio row

Gross margin is operating result over sales revenue - a line the statement
reports but does not add up - and it needed a third `Row.kind` rather than a
flag, because three separate things follow from it: it is drawn italic, its
absolute variance is in percentage *points* ("+8.6%p", not "+8.6"), and it takes
no part in the walk. A margin that advanced the running total would have the
statement adding a percentage to a revenue.

`waterfall_spans` skips it, and its back-up loop now *counts* rows rather than
indexing them, because a ratio row sitting inside a spans-n range occupies a
position without contributing anything. `subtotal_formula` skips it the same way.

It is derived, not typed - `Template.ratio_of` names the two rows it is the
quotient of, and both renderers read it - which makes it the one row on the sheet
that is neither an input nor a sum, and the one where a plausible static value
would survive every other check. `test_responsive` retypes a revenue line and
requires the margin to move, because both of its terms did.

**T04A draws no variance on it at all.** A drawn variance would be a bar in the
column's unit and a margin is not in that unit, so rather than mix denominations
inside one panel the reference leaves the row empty. The series carry None there
and both engines already skip a None.

### One decimal over more precise data

T03A and T04A print to a tenth and compute on the full figures, so the walk lands
0.1 off some subtotals - once per scenario, then carried down every result line
below it. Eleven of T03A's eighteen subtotals sit there, and the tie-out asserts
the *count* rather than hiding it behind a tolerance: a transcription error large
enough to matter cannot then sit inside a number chosen for comfort.

T04A is further out. Its relative variances come back within about a quarter of a
point rather than a tenth - Maintenance is 5.6 over 22.0, which is 25.45%, where
the reference prints 25.3, so the underlying figures are nearer 5.58 over 22.04.
The tie-out states the worst case and names the row.

This is where the standing rule earns its keep. **The SVG prints what the
original printed** - `Tier.printed` carries the reference's labels, so the page
shows +37.3 where our own data gives +37.4 - **and the workbook prints what its
own data says**, because a live sheet has nowhere to put an override. The two
deliverables disagree by a rounding, on purpose, and each is internally
consistent.

It also forced a correction to `verify_table`. It had been comparing the sheet
against the *transcribed* values, which flagged that legitimate 0.1 as an error
on every build. It now compares against `workbook_values()` - the same walk the
sheet's own formulas do - which tests what is actually in question, whether the
formulas do what they claim. That the transcription matches the reference is
`check_ties`' job, and a different question.

### An error in the source

T03A's Group result ΔPY is **wrong in the published reference**. The row reads
236.1 prior year and 422.3 actual, so ΔPY is +186.2 and +79%; the reference
prints +162.5 and +69%, which are the ΔPL figures repeated with the percentage
then taken from the wrong number - so the two agree with each other and with
nothing else. It contradicts the same table twice: the row's own three figures,
and the walk (ΔPY of the result after tax less ΔPY of the profit to other
investors is 192.9 - 6.7 = 186.2).

This is the first outright error the project has found, as against the roundings,
and it is the one place the "print what the original printed" rule does not
apply: reproducing it would put a number on the page that the same page's own
figures contradict, and a recreation that disagrees with itself is the thing this
project has consistently refused. So the derived value is drawn, the printed
label for that one cell is blanked to None, and `_check_t03a` asserts the
discrepancy - if a later pass "corrects" the transcription to match the
reference, it fails and says why.

### Impact is per row, and one series cannot hold two colours

The bug worth recording, because the sheet looked entirely plausible with it:
T04A's panels coloured every cost increase green. The split series were keyed on
the *sign* of the variance, which is right for a table of countries where every
row is a revenue, and wrong for a statement, where a cost line up is adverse and
a cost line down is favourable - so within one column the favourable bars sit on
both sides of the axis.

The fix is to split by impact rather than by sign: the "good" column takes the
value where `higher_is_better_at` says the direction is desirable and NA()
otherwise, and the comparison in each cell's formula is chosen by that row. Two
series still, not four - what varies is per row, so it belongs in the formula
rather than in more series. The live claim raises amortization above plan and
requires the bar to cross the zero rule *and* turn red.

### Small mechanics worth keeping

* **`= Sales revenue` is a formula to Excel.** The label column is formatted "@"
  before the labels are written, or the result lines show #NAME?. Already
  recorded for chart sheets; it bites on a statement table too.
* **A group subtotal states a magnitude.** Operating expenses is 565.2, with the
  minus carried by the "-" against the label, so the formula is `=ABS(...)` of
  the signed sum - which is the `abs(hi - lo)` in `waterfall_spans`, said in a
  formula so the two cannot disagree about the sign.
* **A pair is headed once, at the far edge of the pair.** ΔPY heads both the
  absolute and the percentage column, right-aligned to the end of the second.
  Derived rather than declared: an unlabelled column measured against the same
  scenario is the second half of a pair. That also fixed a caption that read
  "AC" - a variance column falling back to its series' scenario name, which is
  right for a measure column and wrong for every other kind.
* **The label column geometry is three numbers, not one.** The prefix sits at
  the margin, the label 16px right of it, and a component line 9px right of
  that. The margin row looks indented only because it has no prefix.

### Where the table sheets diverge from the reference

* **No spacer rows between groups.** The reference leaves a gap between
  continents; the sheet runs its rows contiguously. This is also what makes the
  panel alignment exact - reproducing the gap would mean blank categories in
  every panel to keep chart and cells in step, and a blank category is a row the
  chart draws nothing on rather than a gap between rows. Recorded as a
  divergence, not defended as an improvement.
* **A panel column's cells are formatted `;;;`.** The panel draws the figure, so
  the cell must not also print it under a transparent chart. The formula stays,
  which is where anyone tracing a bar back to its arithmetic will look.

## Structure charts

C01 and C02 are the first templates whose stack bands are *categories* rather
than scenarios, and that one difference decides the rest of their design.

### A band is not a series

Every stacked thing built before these was a stack of scenarios or of waterfall
steps, so `Series.scenario` carried the meaning. A business area is not a
scenario - every band on C01 is an actual - so reusing `Series` would have put
"Software" in the field whose whole job is to say AC or PL, leaving the notation
with nothing to carry the scenario.

Hence `Segment`, and `StructurePanel` to hold a panel of them. The scenario moves
to the *column*: C01 is four solid actuals and an outlined plan, and the plan
column is what the notation reads, not any band inside it.

### One scale, three panels

C01's whole argument is that 101.9 is drawn the same height in the time series,
in the industry breakdown and in the region breakdown, so a reader can carry a
total across. Excel's default is the opposite - each chart to its own tallest
column - which would draw three quite different numbers at identical heights.

So the maximum is computed once from the tallest column on the sheet and set on
every panel, and `verify_structure` asserts they still agree. It is not a
tidiness check: three panels on three scales is a chart that lies.

The cost is that the maximum is fixed at build time, so a large enough edit
clips a bar rather than rescaling. `test_responsive` checks for exactly that and
says so, which is the right way round - a clipped bar that announces itself
beats a silent rescale that breaks the comparison.

### The plan column, and a divergence recorded twice

C01's 2026 column is outlined *and* its base band is drawn near-white where
every other column draws it darkest. Read as a rule: the darkest band is what
makes a column look solid, and a solid column reads as measured, so the band
that would carry that weight goes to the plan fill instead. Stated as an
interpretation, not as something the standard says.

The reference uses `#F2F2F2` for it - the same value C05X used, and the second
render in the library to prefer it over the house `#FFFFFF`. Two independent
sightings now; recorded rather than adopted, per the palette policy.

### The shade order is ours

Neither reference ramps its greys. C01's business-area panel runs dark, light,
mid, lighter, mid-dark from the axis up, and C02 runs black, light, mid - both
assigned so that no two touching bands are alike, and neither in any order a
function could produce. That is a choice about one chart rather than a rule.

This skill's position is that a caller asks for "the third structure colour" and
never for a hex value, so `structure_colour` is used and the shade *order*
diverges. Geometry, totals, notation and label suppression are reproduced
exactly; the greys are ours. It is the most visible divergence in the library
and the one most worth stating plainly.

### Suppressing a label is a feature

Both templates leave a figure off a band too thin to hold one, and both list it
under required features. The SVG tests the band in pixels and the workbook tests
its share of the axis; on C01 the two agree exactly, dropping 0.9 and 2.9 while
keeping 3.9.

Where the reference suppresses, the value still has to come from somewhere. C01's
two are recovered from the column totals and C02's five from the channel
subtotals - and in both cases the recovery is safe *because* it is over-
determined: C02's three channels each have to add to a figure the reference
prints twice, for Europe and for the World, and all three land exactly.

### The totals row as the legend

C02's Europe and World rows are not bars. They are a swatch and a figure per
channel, then the total and its share - which is this template's integrated
legend, and the reason there is no legend box on the page.

In Excel they cannot be rows of the chart at all: a bar chart draws bars, and a
bar for Europe would be six times the longest country and take the scale with
it. So the chart plots the nineteen element rows through a multi-area range and
the legend rows stay on the sheet as cells.

**They are also where a typed total hides.** They are the only rows on the sheet
the chart does not draw, so a stale one changes nothing about the picture. The
live claim caught exactly that: Austria's total followed an edit while Europe's
and the World's did not, because the subtotal *columns* were still being written
as values. Same rule as everywhere else in this project, transposed - a total
with components is a formula - and the check that found it was written before
the bug was known to exist.

### Excel mechanics from these two

* **A period sum needs a line series, not another band.** A stacked series
  cannot put a label outside itself, so a total carried as a band lands on top
  of the band below it. A line series is exempt, its points sit at the column
  tops because that is what the values are, and its labels go Above. On a *bar*
  chart the same trick is unavailable - Excel refuses to mix bar and line, the
  same limit that costs the pin its head - so C02's row totals ride a
  zero-width band and sit centred on the bar end. On the narrowest rows that
  clips the last band's figure; the SVG carries the faithful version.
* **A data label can be linked to a cell**, with `DataLabel.Formula`, which is
  what keeps a total live rather than written once. The sheet name must be
  quoted: a sheet called `C01A` is itself a valid A1 reference, so
  `="C01A!$C$13"` fails with a bare E_FAIL and `="'C01A'!$C$13"` succeeds.
* **Set a chart title after `strip_chrome`, not before.** It switches the title
  off, and re-enabling afterwards gets Excel's "Chart Title" placeholder back.
* **A transparent chart area exports as black.** Fine on the sheet, and it makes
  every exported check of the built workbook a lie.

## Lines

C07 is the first template drawn as lines, and the first with four series sharing
one pair of axes: a moving annual total across the top, and three cumulatives -
plan, actual to August, forecast after - climbing to meet it, over a tier of the
monthly figures they are the running sums of.

### The marker is the notation

IBCS requires a line chart to show its markers, and the reason is not decoration:
the line between two points is a connector, not data. There is no value at half
past March, and a line without markers invites a reader to read one.

The marker is also where the scenario lives, because a line has no fill. Solid
for an actual, hollow for a plan, hatched for a forecast - so one line can change
from measured to expected part way along, which is exactly what happens in
September. `scenario_marker` says it with the same palette a bar uses, which is
what lets a reader who has learned the fills read a line chart untaught.

### Transcription: the hardest so far, and why

The reference labels only where a label fits, so the printed figures are a subset
scattered across four lines, and two of them can be read against the wrong
series. At November the plan is 1 902 and the forecast 1 921; swapping them is
invisible on the page and wrong everywhere else.

Reading settled nothing. Measuring did: the plan columns at November and December
are 107 and 86 units tall, which only works if the plan cumulative runs 1 794,
1 902, 1 990. The other assignment needs 127 and 69. Everything else then follows
from three sources that have to agree - the monthly columns, the cumulatives they
sum to, and the two bracketed variances - and `check_ties` holds all three
together.

**The structural check is the last point on the page.** A moving annual total in
December *is* the year, so MAT has to land on the forecast cumulative: two series
transcribed from different parts of the chart, meeting at one value.

### A trap worth naming: `_split` had a default

`_split(values, keep)` masks a series to the periods matching one scenario, and
it took its scenario pattern from C03A - a module-level constant, silently, for
every caller. C03A turns to forecast in October and C07C in September, so C07C's
forecast mask came out a month early: September's 1 544 landed on the actual line
instead of the forecast one.

It renders perfectly. The line has the right shape, the right number of points
and the right labels, all one month out. The pattern is now a parameter, the
default is documented as being for the callers that predate it, and the docstring
says what an off-by-one mask costs.

### One printed list per line

`Tier.printed` is per tier, so a tier holding three series gives all three the
same labels - and C07's three cumulatives are labelled quite differently: the
plan at every month, the actual only from March, the forecast throughout. Drawn
that way each line printed the others' figures.

So C07 has one tier per line rather than one per group. And a None in `printed`
had to change meaning: elsewhere it means "no override, use the derived value",
here it means "the reference prints nothing here, because the label would
collide". The renderer treats it as suppression rather than fallback, which is
the only reading that reproduces a chart whose labels thin out where it gets
crowded.

### Excel: a combo is allowed, a bar combo is not

Excel mixes column and line series in one chart quite happily. It is only *bar*
and line it refuses - the same limit that costs C04A its pin heads and C02A its
outside labels - so a template that stacks downwards loses what a template that
stacks upwards keeps.

Two details from the port:

* **The cumulative rows are formulas**, and the forecast cumulative sums the
  actual months *and* the forecast ones. A forecast line built from forecast
  months alone looks entirely normal and stops being true the moment an actual
  is corrected. The live claim retypes March and requires December's forecast to
  move by the same 60.
* **The forecast line borrows the handover month** so it joins the actual line
  instead of starting in mid-air, and draws no marker there - the point belongs
  to the actual.

**Known limit:** an Excel marker cannot be hatched, so the forecast markers are
outlined rather than striped. The SVG draws them properly.

## Areas, and a stock

C08 is C07's axis with a fill, and a stock measure rather than a flow - which
turns out to matter more than the fill does.

### Three tiers, one scale, and two of them share a zero

12.5px per ton everywhere, measured rather than assumed: nineteen decrease
columns give a median of exactly 12.500, and the inventory line fits the same
figure to within a pixel at nineteen of its twenty-one vertices. The two that
miss are the hatched forecast quarter, where a scan up from the floor stops on a
white stripe of the hatch.

**The level and the two movement tiers share a zero line.** The increase columns
rise into the stock they add to. That is not a saving in ink - it is the
relationship the chart is for, and it decides the Excel port below.

### A stock is drawn at boundaries, a flow inside them

The level has twenty-one points for twenty quarters: an opening balance and
twenty closings, drawn on the quarter boundaries, while the movements are drawn
in the quarters they belong to. A flow and a level cannot share an x position and
still mean anything - which is why the template has three tiers rather than two.

The whole chart follows from two transcribed series and one opening balance:
increase and decrease are read off the bottom tier, the change tier is their
difference, and the level accumulates. Twenty-one printed levels then have to
come back out, and they do, exactly. It is the strictest tie-out in the library
because it is a *recurrence*: one wrong quarter throws every level after it, so a
single error cannot be absorbed anywhere.

### The movement shades run opposite to a waterfall's

C08 draws an increase at #A6A6A6 over a decrease at #595959 - light for what
arrives, dark for what leaves. A statement waterfall does the reverse: the dark
bar is the thing being built and the light one the deduction from it.

Both readings are right for their own chart, so `movement_fill` uses the
palette's two waterfall steps with the sign flipped rather than inventing a
second pair that would say the same thing twice.

The palette gained FC and PL waterfall shades on the way, measured off this
template. It had only AC and PY, which was a real gap: an increase and a decrease
have to be told apart in every scenario, not just the two a P&L happens to use.

### Excel: what an area will and will not do

* **Column and line combine; area and column do not.** Setting one series to
  area converts the whole group, and the columns come out as stacked areas. So
  the level and the flows are two chart objects in the same position - the area
  behind, the movements in front on a transparent ground - with the same bounds
  and the same locked plot, which is what keeps their categories over each
  other. Three charts in total, not the two the shared zero would suggest.
* **An area series carries one fill for the whole series.** A column series can
  be painted point by point; an area cannot, so it cannot change from actual to
  plan part way along. The level is therefore three areas, each holding values
  only where its own scenario runs and NA() elsewhere, with the runs overlapping
  by one quarter so the areas meet rather than leaving a gap.
* **A decrease has to be negative to hang below the axis.** The rows hold
  magnitudes, because magnitudes are what the reference prints, so the chart is
  fed a negated scaffolding row - and the two flow series need Overlap at 100,
  or an increase and a decrease from the same quarter sit side by side instead
  of above and below.
* **A column series sharing a chart with an area series has a data-label
  collection Excel will not take a Position on at all** - not the outside ones,
  not centre. The formatting is applied where it is accepted rather than the
  build being failed over a label's placement.

## The shipped workbook

One workbook, a Read me sheet, then one sheet per template. `ibcs_excel.py
--template C03A,C04A` builds them into a single file.

Per-template layout lives in `scripts/ibcs_layout.py`, not in the renderer: the
column plan, the declared widths, the chart-feed formulas and the tier geometry
are facts about a template, and holding them as module constants in the renderer
is what made C03A the only template that could be built. Headers are derived from
the template rather than written out - the reference scenario comes from
`tier.reference`, so a dPY chart cannot end up with a dPL header - and the
formulas are written with column *keys* that resolve to letters from the plan, so
reordering the plan cannot silently break a formula.

The value-axis bounds stay literal, but `check_scale()` now proves the rule they
encode: the measure tier fixes units-per-point, and a same-unit variance tier's
range must equal its own plot height times that rate. For C03A the range is exact
(84pt at 1.0952 = 92.0, and 55.5 - -36.5 = 92.0); only the split of that range
about zero is a rounding of the reference's 41:27, and it is kept as drawn.

---

## Attribution

The reference renders carry an IBCS Institute copyright footer. Recreations must
carry their own footer, not theirs.

C03A's original message line contains a typo ("prevoious"); the recreation
corrects it deliberately.

---

## The XY pair: C10D bubble chart and C09C scattergram

The first two templates in the library with no category axis. Every other chart
here puts a name on one axis and a number on the other; these put a number on
both, and a third number in a third channel — the size of the mark in C10, a
family of curves behind it in C09.

### Three measures, three units, no subject unit

C10 plots market attractiveness (an index), relative market share (a ratio) and
net sales (money). No one of them can head the chart, so **the subject line
carries no unit at all** and each measure states its own where it is used: two
axis titles and the size legend's heading. `TitleBlock.unit` is empty and both
engines suppress the dangling "in" — the worksheet's subject cell is
`=measure & IF(unit="","", " in " & unit)`.

C09 does the opposite and names two measures on one subject line — **Net sales**
in mUSD, **margin** in % — with the measure words bold and the units not. No rule
recovers which words are bold from a flat string, so `TitleBlock.subject` holds
explicit runs for the page. The worksheet keeps the flat text, because a cell
formula cannot carry a weight.

### Recovering the data from the picture

Neither template prints its values, and an earlier note in this file said C09's
~130 points would therefore have to be **synthesized**. That was wrong, and
usefully so: a value that is not printed is still *drawn*, and a drawn position
is a measurement. All 20 bubbles and all 149 markers were recovered from the
reference renders, and nothing in either template is invented.

The method, which is now the standard one for any mark-based chart:

1. **Classify by hue, not by colour.** The markers are translucent — C09's
   overlaps, its curves and its shaded segment all show through — so a mask
   keyed on the exact fill finds a mark over paper and loses it over anything
   else, which is most of the interesting part of the chart. Drawing orange over
   grey changes how light it is and not what colour it is, so hue survives.
2. **Fit discs, don't measure blobs.** A bounding box is wrong the moment two
   marks touch. Correlating the mask with a disc of the marker's radius gives
   one peak per mark, and non-maximum suppression at the marker radius separates
   touching pairs without splitting single marks.
3. **For an occluded mark, correlate against the union.** A mark hidden behind
   another is still a full disc inside *(its own colour) ∪ (the marks already
   found)*. This is what recovers C10's VAB prior year, which is 32 mEUR of
   bubble with a 29 mEUR bubble sitting on top of it.
4. **Fit the radius too, when it means something.** C10's bubble sizes were
   fitted with centre *and* radius free, so the area law could be checked rather
   than assumed. Regressing the 20 measured radii on √value gives
   `r = 14.995·√v − 1.79px` — a slope of exactly 15 and an intercept that is the
   white outline each bubble carries, which a measurement of the visible disc
   loses at both edges.

**What the measurement is worth.** C10's forty coordinates all land within 0.005
of a two-decimal value; that is not something a sloppy fit produces. The radius
regression fits all twenty marks to 0.71px rms, and misreading a single value —
14.0 as 4.0, say — takes that to 5.7px. It has one blind spot, stated in the
check: 9 against 8 draws radii 2px apart, which is inside the measurement noise.

### C10's fills are area fills, not bar fills

Measured off the reference and recorded in the palette under `bubble`: the
actual is the house AC `#404040` at **80% opacity**, and the prior year is
`#D9D9D9` rather than the house PY `#A6A6A6`.

Both departures carry weight. The opacity is not decoration — the bubbles
overlap on most of the page, and an opaque actual would hide the prior year it
exists to be compared with; VBR's prior-year label reads straight through its
actual in the reference. The lighter prior year is the large-area rule the area
charts already follow (C08H fills an *actual* area at `#D9D9D9` too): a mark
forty times the size of a bar carries forty times the weight at the same fill.

Read them as area-specific values of the same semantic system, not a second
scenario palette. Solid dark is still the actual; only the tone moves with the
size of the mark.

### Paint order is notation

C10 paints prior year, then actual, then the acquisitions. Not biggest-first,
which is the usual bubble-chart habit: VAB's actual (29) is *smaller* than its
prior year (32), so size-ordering would bury this year's position under last
year's. What the reader is being asked about is where each unit stands **now**,
so the actual goes on top and the acquisitions — the subject of the message — go
on top of that.

C09 is the opposite and deliberately so: 149 marks of which none matters more
than its neighbour, so the lines go down in legend order and the translucency
does the work.

### A class that is not a scenario

C10's two December acquisitions are drawn in the highlight blue `#0064FF`, and
that is exactly right: they have **no prior year inside this group**, because
they were bought in December. There is no scenario for "arrived too recently to
compare", and inventing one would have put a fourth grey on the page. The
annotation channel is what a highlight is for, and it can never be mistaken for
a scenario because it is not a grey. `S.bubble_fill` returns it for any key that
is not a scenario at all.

C09's three product lines are a different case again: a *category*, coloured from
the accent ramp and given a legend. IBCS's whole point about scenario fills is
that they need no key — solid dark means measured on every page ever printed. A
product line means nothing outside this chart, so the key is what marks the
colour as carrying a category rather than a scenario.

### Iso-gross-profit curves

C09's third measure is the product of its two axes, so a constant gross profit
is a hyperbola: `net sales = 100 × profit / margin`. Three are drawn (1, 2 and
3 mUSD) and the top one is also the boundary of the shaded segment the message
names — the same curve, not a fourth line, so the reader is never asked to trust
an unlabelled boundary.

Two things worth stating. The curve is **stepped in x from where it leaves the
top of the plot**, not from zero, where it runs to infinity. And it must be
drawn as a curve: a straight line between the same two endpoints sits as much as
a third of the plot away from it in the middle.

In the workbook the curves are the sheet's formulas, and the reason they matter
is that they must *not* move when a product does. A curve is the locus of a
gross profit, not a fit through the points; one that shifted with the data would
mean it had been fitted.

### The reference's own count is one more than its picture

C09's message says "we had **45** products of the product line VA in the gross
profit segment of 3 mUSD and above". The chart draws **44**.

This is not a measurement failure. The detection is stable at 44 across every
threshold, marker radius and occlusion setting tried, and a search for fill that
the detected marks do not explain turned up nothing mark-sized. The explanation
is the honest one and it is the lesson of the template: **two products of the
same line at nearly the same margin and the same net sales draw one exactly on
top of the other, and where the colours match there is no edge left to see.** A
scattergram cannot be counted by looking at it.

So the recreation prints the sentence unaltered — the company had 45 — and the
tie-out asserts the size of the gap rather than papering over it. Inventing a
150th product to sit under an existing one would make the arithmetic agree and
the transcription false. This is the same treatment T03A's published error got.

The workbook makes the point visible rather than merely documented: the count is
a `SUMPRODUCT` over the products, in a labelled cell, directly above a second
cell holding the 45 the message states. A reader sees both.

### C10 suppresses three names; C09 places ten

C10's reference prints 17 names over 20 bubbles. Where a pair sits close enough
that the two names would run into each other it prints one and keeps the prior
year's — which in all three cases is the left-hand bubble, so the pair still
reads left to right. Transcribed as `C10D_UNLABELLED` rather than computed,
because where the reference draws the line is far too fine to make a rule of:
RS6's two names would sit 32px apart and one goes, VBZ's would sit 33px apart
and both stay. Every *value* label is printed; only names give way.

C09 names ten products of 149 — the seven at the profitable corner and the three
at the other end. The bottom three are exactly the three lowest gross profits on
the page, which is a tie-out. The top seven are editorial: five of them are the
five highest gross profits but RX-4 and RX-6 are named over two unnamed products
that earn more, so no rule reproduces the set and it is transcribed.

### Divergences

| What | Reference | Ours | Why |
|---|---|---|---|
| C10 value label on the acquisition blue | black | white | `on_fill` decides label colour from the fill it sits on, and white on `#0064FF` is the more legible of the two (4.93:1 against 4.26:1). Two labels of twenty. |
| C09 shaded segment, in Excel | filled above the 3 mUSD curve | not filled | Excel cannot fill the region above an XY series without a stacked-area scaffold that would take its own axis. The SVG fills it; the workbook draws the curve and the count. |
| C10 name above bubble + value inside, in Excel | two labels per point | one, stacked | An Excel series carries one data label per point. The workbook puts the name over the value in a single centred label and says the same two things. |

### Fidelity, measured

**C10D.** All twenty bubbles land within **3.3px in x, 1.6px in y and 2.2px in
radius** of the original, and eighteen of the twenty within 1.1px. The two worst
— VBZ's prior year and VBR's actual — are precisely the two whose *reference*
measurement had the largest fit residual, because they are the most heavily
overlapped and have the shortest visible arc: the discrepancy is in reading the
original, not in drawing it. Every text block on the page is within **2px**, most
within 1.

**C09C.** All three iso-profit curves land within **1px** of the original at
every sampled margin. Of the 149 marker positions, **144** show their own
product line's colour at the exact centre pixel in both renders; the five that
do not are centres crossed by a neighbouring marker's 1.2px outline, which is a
one-pixel sampling artefact rather than a position error. Every text block is
within **2px** except the axis unit line, corrected to 1.

### Two win32com traps this pair added

**`BubbleSizes` wants an address string, not a Range.** Handing it a Range
object is rejected, and the error names the `Values` property rather than this
one, which sends you looking in the wrong place.

**Writing 149 rows a cell at a time makes Excel refuse calls.** `Call was
rejected by callee` reads like a bug in the calling code; it is the COM queue
backing up. Assigning a whole `Range` at once — `.Value` with a tuple of
one-tuples, or a single relative `.Formula` that Excel fills down — fixes it and
is an order of magnitude faster. Worth doing on any block over about fifty rows.

### One sheet with no formulas, and why that is not a gap

C10D is the only sheet in the workbook that computes nothing. That is a fact
about portfolio charts: three independent measurements, no arithmetic between
them. The standing rule — a total with components is a formula — has nothing to
attach to.

So the sheet's claim is the other half of *live*: that the chart reads the cells.
The check is made on the picture, because the sheet cannot show it. Doubling the
bubble that sets the ruler must halve every other bubble's area **including
bubbles in other series**, which is the part that fails if someone ever tidies
the chart by splitting the three scenarios into three charts — a change that
looks like an improvement, leaves every number on the sheet correct, and quietly
puts each scenario on a ruler of its own. Measured on the exported image, the
actual series covers 43,456px before and 21,645px after: a ratio of 0.50.

C09C, by contrast, has plenty to compute — gross profit per product, the count
in the segment, and the three curves — and every one of them is the kind that
must never be typed.

---

## Tree charts

C11A is six small column charts and the arithmetic between them: return over
net sales is return on sales, net sales over invested capital is capital
turnover, and the two multiply to return on investment. It is the only template
in the library whose subject is a *relationship* rather than a set of figures,
and almost everything below follows from that.

### Three typed rows and three derived ones

The right-hand column of the tree - return, net sales, invested capital - is
typed. The other three boxes are quotients of them, in the data layer and in
the worksheet both. This is not a convenience. A tree draws connectors that
state a calculation, and a sheet whose ratios were typed would go on drawing
those connectors over numbers that no longer stood in that relation: nothing
about the picture would change, it would simply stop being true.

`verify_tree` therefore checks that the three ratio cells hold formulas, and
`check_tree_is_live` doubles net sales and requires return on sales to halve,
capital turnover to double, and **ROI not to move at all** - net sales cancels
out of return/capital, so a sheet that moved ROI would be computing the tree
some other way than the connectors claim. Sabotaged by typing the ratios, the
first check names all three rows and the second names two.

### 2023's return is recovered, not read

Every figure IBCS prints here is rounded to a tenth, and a quotient of two
rounded figures compounds both roundings - the trap C03A's PY derivation
already documents. Typing return 2023 as the -3.5 the box prints makes return
on sales come out -13.3 where the original prints -13.1, and ROI -18.0 where
the original prints -17.8.

The precise value is recoverable, and by two independent routes that agree:
the printed -13.1% over net sales of 26.4 gives -3.4584, and the printed
-17.8% over capital of 19.4 gives -3.4532. Two separate printed ratios,
agreeing to 0.005 kEUR, whose mean rounds to -3.5 - which is what the return
box itself prints. With `-3.4558` in the cell, **all twenty-one derived labels
reproduce IBCS's own print exactly**; without it, two do not.

Worth stating plainly because it is the general method, not a C11 trick: where
a template prints a quotient, the quotient is extra precision about the
operands, and using it is recovery rather than invention.

### The scale belongs to the unit

C11's certification requirement is that charts measuring the same thing are
drawn at the same scale. Three groups: the two percentages, the three currency
boxes, and the dimensionless ratio on its own. The scales were **fitted, not
declared** - one number per group, least-squares against the drawn bars - and
the fit is the transcription tie-out: 5.3231 px/%, 4.2793 px/kEUR and 133.4735
px per turn reproduce all **42** measured bar extents with a worst residual of
0.84px. A per-box fit would have been the easy mistake, and this is what rules
it out.

**The box heights differ and that is the consequence, not a contradiction.**
Return peaks at 5.5 against net sales' 27.7 and both are drawn at the same
points per kEUR, so return's box is a third the height. Rescaling it to fill
its own box is exactly what the shared scale exists to prevent.

`verify_tree` states this as a *height* rather than as a ratio: take the
largest value in a group and ask how many points it draws in each box of that
group, and fail if two disagree by more than a point. Comparing rounded
points-per-unit instead fails on a third decimal no eye could resolve, which
tests the arithmetic rather than the picture.

### A negative column takes the lighter tone

Measured, not assumed: the positive columns are `#404040` and the negative ones
`#7F7F7F` - the same step of the ramp a waterfall gives a subtracting row, and
for the same reason. Within one box a bar that adds to the measure and one that
takes away from it have to be told apart, they are the same scenario, and one
scenario fill cannot say two things. `S.waterfall_fill(scenario, -1)` already
had the shade.

### Excel mechanics from this template

Five of these cost real time and each fails by drawing something wrong rather
than by raising.

**A visible category axis caps `PlotArea.InsideHeight`, silently.** Ask a 166pt
chart for a 112pt plot with tick labels under it and the assignment neither
raises nor takes: it comes back 71pt with `InsideTop` driven to *-4*. Excel is
reserving whatever the labels need. Re-locking does not help, because nothing
failed. `_fit_tree_plot` grows the chart by the shortfall and re-measures until
the plot really is the height the shared scale needs - the same move the table
sheets make to align a chart to a row band. Accepting the clamped height would
put every box on its own ruler and quietly undo the whole template.

Because the boxes grow, **nothing can be stacked until every box has been
fitted**. `build_tree_sheet` runs in passes: build and fit, then position from
the heights they actually came out at, then re-assert, then draw every shape
from the finished plot geometry.

**`xlLabelPositionAbove` is a line and XY position.** On a column chart it
raises a bare "Exception occurred" naming nothing. `xlLabelPositionOutsideEnd`
is the one that means "past the end of the bar", and it already follows the
sign - over a positive column, under a negative one - which saves positioning
any point by hand.

**A number format applied before a formula stores the formula as text.** Set
`NumberFormat = "@"` and then `.Formula`, and the category axis reads out
`=B10` in place of the year. It renders perfectly and is nonsense. Format
afterwards, or not at all.

**A multi-level category axis centres its outer group, and draws a separator
grid.** Feeding the axis a two-row range put "AC" under 2023 and "PL" between
2026 and 2027 - Excel centres an outer label over the periods it spans, where
the reference puts each under the *first* period it applies to - and boxed
every label in a grid that has no property of its own to switch off. The axis
is now single-level, its line hidden, with the zero rule drawn back as a shape
and the scenarios as two cell-linked text boxes each. Still live: the cells
they point at hold the suppression rule itself.

**Excel's automatic single-series title comes back after a relayout.**
`strip_chrome` removes it; resizing or moving the chart restores it. It has to
be set false again after the fit and after the move.

**`TextRange2.Characters` is not reachable through pywin32** - it answers "Does
not support a collection" - so a mixed-weight line cannot be built inside one
text frame. The bold measure name and the plain unit are two boxes, the first
auto-sized to its text and the second starting where it ends, which measures
the gap rather than estimating it.

### Where C11A diverges from the reference

- **Plan columns are white, not `#F2F2F2`.** C11's reference fills them
  `#F2F2F2` with a black outline. The house palette settled on white after
  C04A, C06F and T04A were found to contain no `#F2F2F2` pixel at all, and one
  outlier does not reopen it. Recorded as a second source inconsistency, beside
  C03A's `#3C3C3C`.
- **Decimals print with a point.** The reference is German throughout
  ("24,9"); the whole library prints "24.9", and C11 follows the library.

### Fidelity, measured

48 drawn elements - six zero lines and 42 bar edges - land within **2px** of
the original, 46 of them within 1px, none over 3px. Measured by edge rather
than by fill, because the two renders disagree about the plan fill and an edge
detector sees both.

### A fourth error in the source

The message reads *"We plan to achieve a ROI of 19,2% in 2024"*. 19.2% is
**2026**: 2024's ROI is 14.7, the highlight oval on IBCS's own page sits over
2026, and the arithmetic confirms it independently. Printed unaltered, with a
tie-out recording the discrepancy - the same treatment as T03A's variance and
C09C's product count.

---

## Small multiples

C13D is fifteen locations of net profit drawn as relative variances from the
average of them all, and it closes the library at seventeen. It is the largest
transcription in the project and, because IBCS published the same data four
ways, the best checked.

### One dataset, four published views, three of them checkable against it

13A and 13B print absolutes, 13C an absolute variance, 13D a relative one. Any
two of those determine the third, so almost nothing here is taken on trust:
**180 absolute values are transcribed and everything else is derived**, then
checked against a figure IBCS printed somewhere else.

The strongest of those checks is the roster. 13A splits the tail of the
distribution as `St. Gallen` plus `11 more small locations`; 13B folds both
into `12 more small locations`. Adding the first two has to give the third, and
it does, on **all twelve years**. That is what licenses drawing the same data
two ways - the two rosters are one dataset seen from two places, not two
datasets.

### The average is not of twenty-five locations

The reference panel is titled *"average 25 locations"*. It is not. Dividing the
panel total by 25 reproduces **none** of its twelve printed figures; dividing by
**15** reproduces **all twelve**. The divisor is the number of *panels*, one of
which stands for eleven locations on its own - so eleven of the twenty-five are
counted once between them.

This is only findable by reconciling, and it matters practically as well as
pedantically: it means the reference series can be a **formula**
(`=AVERAGE(the fifteen panels)`) rather than a sixteenth transcription, which
is what makes the whole page recalculate when a location is edited.

### 13C and 13D relabel the tail panel

Their tail is titled `12 more small locations` and carries the **eleven**-location
series. 13B's tail, under the same title, carries the genuine twelve-location
sum: it prints 51 in 2017, where 13C and 13D print -25%, which is the eleven-
location panel's 45 against the average of 60. Same title, different series.
Printed as IBCS printed it, with a tie-out recording it - the fourth published
discrepancy in the library, after T03A's variance, C09C's product count and
C11A's message year.

### A discrepancy that turned out to be ours

Linz 2018 looked like a fifth. The derivation gave +87 where the reference
prints +88, and it took two separate mistakes to produce it.

The value is 728/832 of the average, which is **87.5% exactly**. Reaching it
through the average - `(v - total/15) / (total/15)` - divides twice and lands
1.4e-14 short of the boundary, so it rounds down. And rounding it with Python's
`round()` lands on 87 again, this time because 87.5 goes to even where Excel
rounds halves away from zero.

Computed as `(v*15 - total) / total` and rounded with `excel_round`, **all 180
printed variances come back exactly**. Two independent ways to manufacture a
discrepancy in a source that did not have one, and worth keeping: the project's
rule is that a check must fail once before it is believed, and this is the
converse - a check that fails should be doubted before the source is.

### One scale, measured

Every one of the fourteen pin panels is on the same ruler, and it is checked
against the drawing rather than asserted: regressing **167 measured stem
lengths** on the derived variances gives 0.59268px per percentage point with no
residual over half a pixel. (167 and not 168 because Linz 2025 is exactly zero
and draws no stem.)

That single ruler is what Berlin costs the other fourteen. It reaches +343%
where no other panel passes +112%, so Salzburg's whole twelve-year series lives
inside 93px - and the cost is the point. A variance read in one panel means the
same in the next, which is the only reason to put them on one page.

**Berlin's panel is twice the height, and that is the same argument.** Its zero
sits on the *lower* row's zero line, which buys it 243px of headroom where one
row gives 83. Rescaling it instead would have been the easy mistake.

### Two arrangements of one set of series

The SVG follows the reference: fourteen pin panels, Berlin spanning two rows,
the average drawn as columns on a shaded ground, St. Gallen not drawn at all.

The workbook draws a **uniform 4x4 of fifteen pin panels** - every cell the same
size and the same kind, Berlin in one row, St. Gallen in its own right, the
tail panel reverting to the eleven locations it actually carries, and the
sixteenth cell left empty. The reference panel is not drawn: a pin is notation
for a variance, and the baseline's variance from itself is a row of zeros. The
average is stated in the subtitle and lives in the input block as the divisor.

`PanelGrid` and `PanelCell` in the data layer hold both, so neither renderer
invents its own arrangement and the tie-outs can check both.

This is the one place in the library where the two engines draw different
panels, and it is a deliberate divergence rather than a limitation of either:
the user's call, taken because a uniform grid of one kind is what a general
panel engine should be asked for.

### Excel: the grid is one chart, and it is not built here

Excel has no panel chart type. The workaround everyone reaches for - one chart
copied sixteen times - leaves sixteen chart objects whose axes drift apart the
first time a figure changes, which is the one thing a small multiple may not
do. So the grid is built by the **`panel-charts` skill**, and C13 is its first
caller: the skill supplies the geometry, this sheet supplies the notation.

What that cost the panel skill is recorded in its own decisions file - the pin
kind, a style seam, cell-linked value labels, an in-panel divider, and a
`build(ws, origin=...)` that lets a caller host the grid on a sheet it already
owns. What it bought is that the next grid of charts anyone needs is already
built.

**The workbook stays live across the seam.** The engine is handed values so it
can size its shared scale, and then every cell of its input block is
*overwritten with a formula* reading the absolutes above it.
`check_panels_are_live` raises Salzburg by 900 and requires three things at
once: the average moves, because it is derived; Salzburg's pin crosses zero and
turns from red to green, because the split is by impact and not by sign; and
**Cologne moves too**, from +31% to -34%, without being touched - because every
panel is measured against the average that just rose. A grid where only the
edited panel moved would be fifteen charts sharing a page.

### Fidelity, measured

181 drawn elements - fourteen zero lines and every pin stem - land within
**1px** of the original, all 181 of them. Measured with the same detector run
over both renders.

### One divergence in the SVG worth naming

The grey tick stubs marking the labelled years sit at the same x as a pin, and
drawing them after the pins tinted the stems - a 1.4px grey line over a 3px red
stem takes it from `#FF0000` to `#E2A3A3`. The stem colour is the only channel
saying which way a variance went, so the ticks are drawn first and the pin over
them, which is what the reference does.

---

## Making a sheet follow somebody else's data

The library was built as a fidelity proof of IBCS®'s own example figures, and
its axis bounds were measured for those figures. Paste your own over the
placeholders and the bars clip or shrink. Fixing that is what turns the
workbook from a recreation into a tool.

Excel will not bind an axis bound to a formula. So rather than scaling the axis
to the data, scale the data to a fixed axis: a hidden **span cell** per scale
group holding the largest magnitude in it, a **scaled copy** of every drawn
column (`= raw / span`), and axis bounds that are the declared bounds divided by
the same span.

Dividing rather than min-max normalising matters. **Zero divided by anything is
still zero**, so the zero rule, the reference rules, the tier captions and every
shape positioned from plot geometry land exactly where they always did. Nothing
in the drawing code moves.

At the shipped data the change is a visual no-op, and that is checkable rather
than hopeful: `axis bound x span` must reproduce the bound the layout declares.
It does, to six decimal places, on every converted tier. The C04A tier images
came back pixel-identical.

**Data labels have to be linked to text cells.** A series plotting a fraction
would print the fraction, and a cell-linked label ignores its own number format
and prints full precision - so the formatting has to happen in the cell, via
`TEXT()`. `scale_columns()` emits those text copies.

### The tiers this cannot work for, and why

Normalising fails whenever a template's declared range is deliberately
**narrower** than its data - which is exactly what an outlier-marking tier is.

The axis is fixed at build time. A series plots `v / span`, and the largest
value in any dataset plots at exactly `1.0`, because it *is* the span. So the
tier fits only if its upper bound, as a fraction, is at least 1.0. C12A's
relative tier stops at +265% over data reaching +983%, a fraction of 0.2695.
Normalising it would make the largest bar run off the panel **whatever the
figures were** - as shipped, with the outliers removed, or scaled down to
single digits. The arithmetic is in the plan; the conclusion is that following
the data and clipping outliers are contradictory jobs, and that tier's job is
the second one.

So C12A's relative tier keeps its declared bounds and clips instead - see
below. The same reasoning, for a different reason, keeps C09C and C10D off the
scheme: they are the only charts in the library that **show** their value axis,
with gridlines and tick labels. `strip_chrome` takes the axis off everything
else, which is precisely why dividing by a span is invisible there. Divide on an
XY chart and the reader sees `0.25` where the data says `29.16%`.

## Clipping an outlier in Excel, and keeping it live

IBCS clips an element that leaves its panel and prints one to three triangles
where it left, the count rising with the size of the overrun, so a reader can
tell a bar that just missed from one that ran off the page.

Three routes were tried and rejected, each of which looked workable:

| Route | What actually happened |
|---|---|
| A line or scatter series carrying triangle markers | **Accepted** onto a bar chart, then drawn against its own axes - the marker landed a row out and brought a secondary value axis with it |
| Custom, range-driven error bars | **Refused outright** on a bar chart, in every argument form - Range, address string and value list alike. `HasErrorBars = True` alone is accepted, which is what makes this look possible |
| Triangles as drawn shapes | Works, and is wrong - for the same reason the horizontal pin has no head. A static ornament beside a bar that has since changed states something false, which is worse than an absent ornament |

What works is the trick the sheet already uses for its variance numbers: the
label. The bar reads a **clipped** copy of its column, so it stops inside the
panel; the label is linked to a **text** cell that carries the measured value
*and* the triangles. One label rather than two, because a separate marker series
lands in exactly the same place and the two print on top of each other.

```
=IF(ISNA(F13),"",
  IF(F13>265, TEXT(F13,"+0;-0")&REPT(UNICHAR(9658),MIN(3,MAX(1,INT(F13/265)))),
  IF(F13<-35, REPT(UNICHAR(9668),MIN(3,MAX(1,INT(-F13/35))))&TEXT(F13,"+0;-0"),
  TEXT(F13,"+0;-0"))))
```

`UNICHAR`, not `CHAR`: `CHAR` stops at 255 and the triangles live at U+25BA and
U+25C4. The triangles follow the number going right and precede it going left,
so they always point away from the axis. Every part of it is a formula, so a
figure edited back inside the panel loses its marker without anything being
redrawn. It reproduces IBCS®'s own counts on the shipped data: three on +983%,
one on +391%.

The label needs `WordWrap` off. "+983" plus three triangles is wider than the
slot Excel sizes the label to, and it wraps - putting the markers on a second
line where they read as a separate element rather than as a continuation of the
bar. Same failure as the tier captions, same fix.

### What each family needed

The mechanism is one idea, but the sheets are genuinely different shapes and
each needed its own plumbing.

| Family | How it scales |
|---|---|
| **Tier stack** C03A C04A C05X C06F C12A | Layout only. `scale_columns()` appended to the column plan, `scale_group` on each `TierSpec`. Every builder already fed through `_feed`, so nothing in the drawing code moved - except `_pin_series_h`, the one series builder still reading its columns directly, and C05X's measure labels, which were hand-rolled instead of going through the helper that links them |
| **Tables** T02A T04A | The feed columns behind the drawn panels divide by a span parked in the same zone; the invisible labeller plots the scaled copy and links to a text one. T01B and T03A print their variances as numbers and draw no bars, so they have nothing to scale |
| **Structure** C01A C02A | The segments lie in *rows*, so the scale block is a mirror of the block written below it and hidden. The span is taken over the drawn columns only - a horizontal panel's subtotal columns are the integrated legend and no bar carries them, so a span over them would be set by C02A's World total of 3,733 where the tallest country is 496 |
| **Line** C07C C08H | Rows again, and one span across the columns and the lines together - which is what lets a monthly bar and a cumulative line sit in one chart and be read against each other. C08H's level and its flows share one span for the same reason: the columns rise into the stock they add to |
| **Tree** C11A | Three spans, one per unit. The box *height* still comes from the raw range, so a box whose numbers span more is taller - that is how boxes sharing a unit come out sharing a scale, and only the axis is divided |
| **Panel grid** C13D | Already normalised by construction - it is the engine the rest of this borrows from |
| **XY** C09C C10D | Not normalised. Fitted at build, keeping the step the template chose |

Two of these changed what a check was asserting rather than what the code did.
`verify_tree` and the tree's liveness check both measured points per unit off
the axis, which now runs in fractions; and the tree check asserted that the
scales **did not** follow the data, which was true when the tree was fitted
once and never again. What still has to hold - the only thing a reader can be
misled by - is that boxes of one unit agree with each other. Doubling net sales
takes all three currency boxes from 3.209 to 2.005 points per kEUR: the same
number in all three, so a figure carried between them still reads the same size.

`scripts/test_rescale.py` is the gate. It multiplies every typed input by a
hundred and asks both questions of every sheet, and it has been made to fail
before being believed - reverting one axis to its declared bound is caught as
both a moved picture and a broken shared scale.

---

## What building on somebody else's numbers exposed

The seventeen recreations prove the notation against IBCS®'s own example data.
Rendering nine charts from a real set of financial statements proves the
*renderers*, which is a different claim - and it found three places where a
renderer was still reading IBCS®'s data rather than the template it was handed.
All three are the same mistake, and it is the one this library is most prone
to: a value transcribed from the reference, used for geometry, and correct
right up until somebody else's figures arrive.

| Where | What it did |
|---|---|
| The **"AC" legend caption** | Levelled itself with `D.C03A_MEASURE[0]`, a transcribed January value. On cash in the tens it floated halfway up the chart. Now takes the template's own first actual |
| The **annotation oval** | Looked its value up in `D.C03A_MEASURE[index]`. Now derives it, and by the period's *own* scenario - C03A's oval sits on a December forecast, where the actual series holds nothing |
| **Measure labels** | Went through `_thousands`, which fixes the decimals at none. Right for kEUR in the hundreds, wrong for cash in the tens: 4.295 printed as "4". Now routed through the tier's own number format, then given IBCS®'s thin-space separator |

Two pieces of C03A's furniture were also unconditional and are now asked for:

* the **forecast rule** and its "FC" caption, which on a chart of three
  measured years drew a rule past the last category and named a scenario the
  chart does not carry;
* the **scale-break index band**, which exists only to make a detached annual
  block honest. Without one it is decoration, and at a three-category pitch it
  is decoration in the wrong places.

The recreation is byte-identical after all five changes, which is the only
reason to believe none of them was a fidelity regression.

### Fitting a drawn variance column

`fit_panel` in the financial-statements builder derives where a table panel's
zero sits and how long a unit is. The part worth recording is that it reserves
**label room**: scaling the bars to the full column and letting the numbers
fall where they may put the revenue variance's "-6.3" on top of the AC column.
IBCS's own T04A puts zero at 97 of 240 because their statement overspends more
than it underspends; this company's does the opposite, so the zero moves.

### The tie-out that matters when nothing is transcribed

`check_ties()` runs before any chart is drawn and compares every derived
subtotal to the number the statement itself prints. With no transcription the
risk is not a typo - it is reading the wrong row, or summing a block that does
not mean what it looks like. Pointing one check at "Operating Expenses Total"
instead of "Total Expenses" reports `built 257.17, statement says 146.80`,
which is how it was confirmed to work.

## A document the workbook cannot disagree with

`ibcs_doc.py` writes the manual-build markdown for a workbook by **opening the
`.xlsx` that was just written and reading the formulas back out of it**. The
technique is borrowed from `panel-charts/scripts/panel_doc.py`, and the reason
for it is simple: prose has no tests. A hand-written guide drifts the first time
a formula changes and nothing catches it, so the guide is generated from the
artefact instead of written beside it.

Three decisions inside it earned their place.

**Formulas are grouped by shape, not listed.** A column of twelve variance
formulas is one idea, not twelve, so every formula is normalised to R1C1 -
`=D13-C13` in E13 becomes `=RC[-1]-RC[-2]` - and cells that agree after
normalisation are reported as one row with the range they cover and one worked
example in ordinary A1. C03A's 281 formula cells come out as 25 shapes; C13D's
2,124 as 101. Without that the second half of the document is unreadable.

**A group is named for every column it spans.** Two scaled copies dividing by
the same span cell say the same thing about their own position, so they group -
and naming the group after whichever column came first labelled `Q13:R24` "ref
scaled" when half of it is the measure. It now lists every distinct header in
the group.

**A one-off formula with a name is not a one-off.** The span cells are a single
formula each and are the most interesting cells on the sheet; filing them under
"5 one-off formulas: B7, B9, B10, O13, P13" buried them. A group earns a table
row if it repeats *or* if the sheet gave it a header.

### The check that runs the other way

Generating a document stops it being written wrong. It does not stop it being
*edited* wrong afterwards, or shipped beside a workbook that has since been
rebuilt - and a document quoting a formula the file does not contain is worse
than no document, because it will be believed.

So `ibcs_doc.py --check` parses the existing markdown, extracts every quoted
formula together with the address it names, and asserts the workbook holds
exactly that there. 360 quotations in the complex document, 182 in the simple
one. Per this project's rule it was made to fail before being believed: two
quotations were hand-edited, one in a technique section and one in a per-sheet
table, so that both of the checker's two parsing paths were exercised, and it
reported

```
C03A!E13: document says '=D13-C14', workbook holds '=D13-C13'
C03A!I13: document says '=IF($E13>=0,$E13,NA())', workbook holds '=IF($E13>0,$E13,NA())'
```

### Two things the document says that were derived, not assumed

Both started as sentences that read well and were wrong.

- **"The chart zone is to the right of the data"** is true of eleven sheets. The
  four tables print the *grid itself*, starting below the six typed title rows -
  the same rule, since the print area holds the deliverable and the typed inputs
  are not it - and C13D parks its charts below, because the panel engine owns
  the columns to the right. The document now derives which case a sheet is in
  from its family rather than describing the majority and hoping.
- **A first-letter test for "starts in column A"** matched `AD1`. Column letters
  have to be parsed, not sampled.

## The simple workbook's bars had collapsed, and every check passed

Found by the user on review, in the four sheets `C04A`, `C05X`, `C06F` and
`C12A`: the labels, the axis and the category names were all correct, and the
bars were drawn at a fraction of a pixel.

**The cause is one missing field.** A simple layout redeclares its tiers from
scratch rather than filtering the full ones - which is right, because a
base-tier sheet gets its own geometry - but that means every per-tier fact has
to be restated, and `scale_group` was restated for `C03A` only. C03A was the
pilot when the scale block went in; the other four were added to `LAYOUTS` and
never to `SIMPLE_LAYOUTS`.

**Why it failed silently, and this is the part worth keeping.** The scaled
columns are written by the *column* declarations, which the simple layout
inherits unchanged, so the values reaching the chart were divided by the span
cell exactly as intended. The axis bounds come from the *tier*, and
`_scaled_bounds()` returns the declared bounds untouched when a tier has no
scale group. So the sheet plotted values around 1.0 against an axis running to
240, or to 2,200, or to 1,120.

Nothing raised. The build's own `verify()` checks tier alignment and zone
separation, and both were fine - the charts were in exactly the right places,
drawing almost nothing. Page setup was fine. The generated document was fine;
it quotes cell formulas, and the cell formulas were correct. Only a human
looking at the picture, or a check that multiplies an axis bound by its span,
could see it.

### Why the gate did not run

`test_rescale.py` is precisely the check that would have caught this - its first
assertion is `axis bound x span == the declared bound` - and it could not be
pointed at the simple workbook at all. It read tiers from `L.LAYOUTS`, so on a
sheet with one tier instead of three it asked Excel for a chart object that does
not exist and died on a COM error. The same was true of the structure check,
which asked for three panels where the simple `C01A` draws one.

A gate that cannot open a file is not a gate that passes on it. Three changes:

- `is_simple_book(wb)` asks the **Read me sheet** whether this is the base-tier
  workbook, rather than trusting the file name - a copy renamed for review is
  still the same workbook;
- tier and structure checks read what the sheet actually holds (`layout_for`,
  and the panel list taken from the sheet's chart objects), and print how many
  of the expected panels were drawn rather than silently checking fewer;
- families absent from a workbook print `not in this workbook`, never a pass.

Run against the pre-fix file it reports nine failures across exactly the four
sheets the user named, and nothing on `C03A` or the structure sheets - which is
the proof that the fix and the gate describe the same defect.

### And a check so the field cannot go missing again

`check_simple_layouts()` in `ibcs_layout.py` asserts that every simple tier sits
in the same scale group as its full twin. It lives in the layout module rather
than in a renderer because it is a fact about the registries, and **both**
renderers call it - `check_registries()` in the SVG path, and `build()` in the
Excel path, which is where the defect actually lived and which had never run a
registry check at all. Made to fail first by putting `C06F`'s `wf` tier back the
way it was:

```
C06F simple tier 'wf' is in scale group None but the full layout puts it in
'unit'; the axis bounds and the values would be divided by different things
```

**The general lesson.** Where one structure is derived from another by
restatement rather than by transformation, the risk is not a wrong value - it is
an *absent* one, which reads as a default. Every field so restated wants either
a transformation that carries it or a check that compares the two.

## The build document became a build guide

The first version of `ibcs_doc.py` was an inventory: every formula on every
sheet, grouped by shape, with its address. Accurate, checkable, and not what
anybody needs. The user's verdict was exact - *"nothing equivalent exists to
help an end user even have a starting point to create something manually"* -
and the benchmark they pointed at was their own `Vertical waterfall - manual
steps.md`: which menu item, what to type where, what you should see afterwards,
and what it looks like when it has gone wrong.

The rewrite keeps the guarantee and changes the genre. It is now in three parts:
the **sheet skeleton** in seven steps (the title block, the two zones, the data
columns, splitting a series by meaning, the scale block, text columns, page
setup); a **worked build per family**, with the tier stack done at full depth -
insert path, `Ctrl+1`, gap width, overlap, hex fills, the pin's error-bar
construction, plot-area alignment - and the other six written as *what differs*;
and a **per-sheet reference** for rebuilding one template rather than learning
the method.

**What changed technically is that the guide now reads the drawing, not just
the cells.** `sheet_chart_facts()` walks every chart object while Excel still
has it open and records chart type, size and position in points, plot-area
geometry, gap width, overlap, axis bounds, and per series the fill, line,
marker, error-bar weight and colour - and, crucially, **how each individual
point is painted**. A guide that says "set the fill to `#404040`" is only worth
having if that hex came out of the file.

### It found two things the cell formulas could not

**An Excel marker cannot take a pattern fill, and says nothing about it.**
`apply_scenario()` called `Fill.Patterned(...)` on the pin heads on every build,
so the forecast months of C03A and C05X were *meant* to be hatched in the
relative tier. They were not. The call is accepted without error, `Fill.Type`
reads back as `-2` (mixed), and the head draws exactly as it did before -
calibrated directly rather than inferred, because a silent no-op is not
something to guess at. The SVG, which has no such limit, hatches them, so the
two engines had quietly diverged on a piece of notation the skill's own
documentation calls load-bearing: *the head is what says whether a relative
variance was measured or forecast*.

The fix is a **hollow head**: the forecast's own ground colour `#F2F2F2` with a
`#404040` border. That is not a compromise outside the notation - it is how this
notation already says "not measured" for plan and budget, and the `panel-charts`
skill had reached the same answer independently when C13 needed it. The
divergence from the SVG, which still hatches, is deliberate and is recorded
rather than hidden.

**A marker series' colours are on the marker, not on the series.** The first
version of the fact collector read `series.MarkerBackgroundColor` and reported
`#156082` - Excel's default accent blue - because the pin heads are coloured
per point and the series-level property was never touched. Reading it put a hex
in the guide that appears nowhere in the file. Per-point reading fixed it, and
it is the same lesson as the hatch: on these charts, the series is not where the
notation lives.

### And one that was only ever a flaky test

`test_rescale.py` intermittently died on `RPC_E_CALL_REJECTED` against the
seventeen-sheet workbook, and the existing remedy was to restart the whole run
up to three times - which stopped working. The cause is not flakiness:
`Application.Calculate()` **returns before the calculation has finished**, and
the next COM call arrives while Excel is still repainting seventeen sheets of
charts. Polling `Application.CalculationState` until it reads `xlDone` turns an
intermittent failure into a short wait. A test that fails once in three runs
teaches people to re-run it, which is how a real failure gets re-run away.
