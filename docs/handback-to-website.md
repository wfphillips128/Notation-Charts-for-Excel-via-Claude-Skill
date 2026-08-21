# Hand-back to the website project

Written for the project behind `edgewisedata.com/charts/notation-charts`, which
asked this one for a second workbook in
[`alternate-data-workbook.md`](alternate-data-workbook.md). That work is done.
Two things the request assumed are now wrong, and both need a decision on your
side rather than ours.

**Nothing here is urgent and nothing here is broken.** The page works. These are
two claims that have stopped being true.

---

## What exists now

Four workbooks in `workbooks/`, two datasets by two forms:

| File | Data |
|---|---|
| `IBCS Charts - Complex Versions.xlsx` | the IBCS® Institute's published example figures, all 17 templates |
| `IBCS Charts - Simple Variants Only.xlsx` | the same figures, base tier only, for the 7 templates that have a useful simple form |
| `IBCS Charts-Progressive - Complex Versions.xlsx` | The Progressive Corporation's reported figures, all 17 templates |
| `IBCS Charts-Progressive - Simple Variants Only.xlsx` | the same, 7 templates |

Each has a matching guide in `guides/`.

Every figure in the Progressive pair is either **filed** (it appears in an SEC
filing or an investor-relations release), **derived** (computed from filed
figures by a rule the sheet prints), or **assumed** (a construction the company
never published — chiefly plan and forecast, which no public company discloses).
Constructed values are shaded in a distinct hue, footnoted on the sheet, and
listed on a **Sources** tab. The harvest reconciles to Progressive's own totals
in 605 automated checks.

---

## 1. `ibcs-charts-placeholder.zip` is the wrong name

The page offers that download and describes it as reference files with
**placeholder data** you can paste over.

That was already untrue. The zip holds the Institute's *actual published example
figures* — Furniture Inc., Berlin, California — because the workbook is a
fidelity recreation, and being a faithful copy is the whole point of it. That
false claim is what prompted the original request.

It is now untrue in a second way: **the thing that fills the gap exists, and it
is not placeholder data either.** It is seventeen templates on a real insurer's
audited filings.

**What to do.** Rename the download and rewrite the sentence. The honest and
stronger offer is something like *"the same seventeen templates on The
Progressive Corporation's reported figures"* rather than *"placeholder data"*.
The `note:` append and the body sentence both need it.

**One constraint if you publish the Progressive workbook.** Its Read me carries
two disclaimers and they need to travel with it wherever it is described:

> IBCS® is a registered trademark of the IBCS Association; this is not an IBCS
> Institute publication and is not endorsed by it. This workbook is not
> affiliated with, endorsed by, or produced in cooperation with The Progressive
> Corporation.

---

## 2. `render-site-svgs.py` carries its own numbers

That script generates the SVG chart images on the page, and it holds
**hand-substituted figures for three templates — C03A, C01A and T04A** — typed
in so the public page would not show Furniture Inc.

That makes three copies of "data for these templates":

| Copy | Where |
|---|---|
| the recreation | `ibcs_data.py` |
| Progressive | `ibcs_data_alt.py` |
| three templates' worth of substituted numbers | inside `render-site-svgs.py` |

Three copies is three chances to drift. It is the same shape as a defect this
project spent a while removing: both renderers used to reach past the `Template`
they were handed and read figures from the data module directly, in 34 places,
and a copy left behind does not error — it quietly draws the wrong numbers under
the right title.

**What to do.** Point `render-site-svgs.py` at `ibcs_data_alt.py` and delete its
inline figures. The seam is ready: `ibcs_data_alt.TEMPLATES` is a dict keyed by
template id, holding the same `Template` objects `ibcs_svg.py` already renders.
Roughly:

```python
import ibcs_data_alt as A
template = A.TEMPLATES["C03A"]      # instead of the local substituted copy
```

**The catch, and it is a real one.** Those three published SVGs would then show
Progressive's figures instead of the substituted ones. That is a visible change
to a live page, so it is a decision rather than a refactor — and it is yours,
not ours. If you would rather not change what the page shows, the inline data
stays and this note is just a record of why it is a risk.

**A gate worth keeping either way.** Before and after any change here, render
all seventeen SVGs and compare: `ibcs_svg.py --template <id> <out.svg>`. The
seventeen from `ibcs_data` should come out **byte-identical**, which is how this
project has verified every engine change it has made. Only the three you
deliberately switch should differ.

---

## 3. Worth knowing: the Excel output changed

Not something you asked about, but it affects any asset regenerated from this
repo.

- **C05X's bars are 2.6× wider.** Its measure tier never applied the cluster its
  own layout declared, so it fell back to Excel's defaults — 150% gaps, no
  overlap — and drew the plan *beside* the actual instead of behind it. Both
  workbooks carry C05X and both changed.
- **C13D's panel scale gained headroom**, so a value label can no longer print
  through the panel title above it.
- Smaller fidelity fixes to C02A's band labels and C10D's suppressed labels.

The SVG engine never had any of these, so **all seventeen SVGs are
byte-identical** to what is published. Only the Excel files differ.

---

## What this project is not doing

Two known gaps, recorded so they are not discovered as surprises:

1. **Excel draws no annotations outside the driver tree.** C07C's two bracket
   callouts and C08H's highlight oval are on the templates and the SVG engine
   draws them; the Excel engine honours annotations only on C11A. This affects
   the recreation exactly as much as the Progressive workbook, and predates
   both.
2. **C09C and C10D fit their axes at build time**, not live on edit — they are
   the only two charts that show a value axis, so they cannot use the
   normalisation every other sheet relies on. Stated in `references/decisions.md`.
