# Building a workbook on data that is not the Institute's

> **Status: SUPERSEDED — this was the plan, and the work shipped on
> 2026-08-20.** Kept as the record of the reasoning, not as instructions.
> For what exists, read the README's *The second dataset* section and
> `ibcs_data_alt.py`.
>
> **What the plan got right.** All three findings held: the renderers really
> did reach past the template they were handed, `ibcs_layout.TEMPLATES` really
> is the same object, and `check_ties()` really is vacuous on substituted data.
> Each became a permanent guard rather than a one-off fix —
> `test_dataset_isolation.py` asserts the *property* (every `ibcs_data` name a
> renderer reads must be a type, helper, registry or `Template`) instead of
> keeping the "exhaustive, auditable list" this proposal called for, so a new
> global fails the day it is written. `test_alt_ties.py` is the second tie-out
> pass, and it runs before `build()`.
>
> **What it got wrong, and it is the useful half.** The data went **real**, so
> every consequence this document files under *if the data is real* came true
> at once — and the tie-out strategy inverting was worth more than expected.
> The target here was ~85–110 internal-consistency checks; what shipped is
> **211 of those plus 605 reconciliations against Progressive's own filed
> totals**, which is a strictly stronger claim than anything invented figures
> could have supported. The estimate below assumed invented data and is
> therefore not a measurement of anything that happened.
>
> **The dataset lives in-repo**, as this document's own table recommends — but
> as a sibling `ibcs_data_alt.py`, not the "parameter, not a second copy"
> refactor of `ibcs_excel`. The isolation test bought the same guarantee for a
> fraction of the cost.
>
> **Attribution went further than proposed.** `READ_ME_OWN_DATA` was added as
> the Excel-side counterpart of `FOOTER_OWN_DATA`, and beyond it
> `Template.provenance` marks each scenario *filed*, *derived* or *assumed*,
> shades an assumed scenario's cells differently, and writes a **Sources**
> sheet. `test_provenance.py` checks both that the marking marks and that a
> template declaring nothing is left exactly as it was.

## Why

This project is a fidelity recreation, and that is the point of it: seventeen
templates transcribed from the IBCS® Institute's published example renderings,
127 tie-outs asserting the labels those renderings print, and pixel tolerances
measured rather than judged. The README is explicit that everything it builds,
it builds from `ibcs_data.py`.

That makes the shipped workbooks a poor fit for one thing people keep wanting
them for: a file to paste their own numbers into. The mechanics are already
there — every sheet is live, the shaded cells are typed inputs, and
`test_rescale.py` asserts that a ×100 paste still plots inside every axis. What
is not there is a set of numbers a reader can overwrite without first working out
which of them are Furniture Inc.'s.

So: a second workbook, same seventeen templates, same renderer, different data.

A consumer already does a small version of this. `scripts/render-site-svgs.py` in
the `edgewisedata` website repo substitutes data for three templates (C03A, C01A,
T04A) and renders them through this project's `ibcs_svg`, precisely so the site
does not publish images of the Institute's figures. It works, and it is the
existence proof that the renderers are data-driven. See *Open question* at the
end — if this project takes on a dataset, that script should probably consume it
rather than carry its own.

---

## What survives a change of data source

| | Section |
|---|---|
| **Source-independent** — holds for invented, real or licensed data | The three findings · Where the dataset lives · Attribution · Tie-out strategy and the three guards · Verification |
| **Source-dependent** — rewrite if the data changes | Per-template notes · "hold magnitudes near the originals" · Effort estimate |

**If the data is real rather than invented, four things change.** All are
expensive to retrofit, so decide before authoring starts.

1. **Licensing and attribution.** Invented figures need only a "this is not a real
   company" disclaimer. Real figures need a provenance line, and possibly a licence
   that has to sit alongside this project's MIT licence without contradicting it.
2. **The magnitude lever is gone.** Invented figures can be authored to fit the
   layouts' fixed scales. Real figures arrive at whatever size they are, which puts
   all the weight on the span-cell scheme and promotes the C11A box-fit and XY
   axis-containment guards from belt-and-braces to load-bearing.
3. **The tie-out strategy inverts.** With invented data you state a minimum and
   derive the rest, so the checks assert internal consistency. With real data the
   checks become a *reconciliation* against the source's own stated totals —
   different, and arguably stronger.
4. **Seventeen templates may not all be expressible.** C09C wants ~70 points across
   three groups with a margin/volume/profit relationship; C13D wants 15 entities ×
   12 periods. A real source either has that shape or it does not, and the fallback
   is a mixed workbook — some sheets real, some invented — which has to say so
   plainly on the Read me sheet.

---

## Three findings, all verified against the code

### 1. Swapping `Template` objects is not enough

`ibcs_excel.py` reaches past the template it was handed and reads `ibcs_data`
module globals in seventeen places:

| Global | Read at |
|---|---|
| `D.C08H_OPENING`, `D.C08H_LEVELS` | `ibcs_excel.py:3289, 3475` |
| `D.C09C_ACCENT`, `_MESSAGE_LINE`, `_SEGMENT`, `_MESSAGE_COUNT`, `_ISO_PROFIT` | `ibcs_excel.py:3698, 3901-3987, 4093` |
| `D.C13D_PANELS`, `_AVERAGE`, `_VARIANCE` | `ibcs_excel.py:4869, 4985, 5037` |

`ibcs_svg.py` reads those plus `D.C10D_ACQUISITION`, `D.C10D_UNLABELLED`,
`D.C09C_LINES` and `D.C11A_SCALE_PX`.

This is the riskiest thing in the job, because it **fails silently**. A global left
behind does not raise — it draws the Institute's numbers on a sheet titled with
somebody else's entity. Whatever mechanism is chosen needs an exhaustive, auditable
list and an assertion that every name it sets already existed.

### 2. `ibcs_layout.TEMPLATES` *is* `ibcs_data.TEMPLATES`

`ibcs_layout.py:1523` is a plain assignment — the same dict object, not a copy. Any
swap must **mutate in place**, never rebind.

### 3. `check_ties()` is vacuous on substituted data

`_CHECKS[template.id](check)` passes only the `check` closure, and the per-template
functions read module globals rather than the `Template` they are nominally
checking. `build()` calls it at `ibcs_excel.py:5404` as a gate, and on substituted
data it **passes while proving nothing**.

A second, real tie-out pass is therefore mandatory, and must run *before* `build()`.

**Corollary, and it is a gift:** `test_rescale.py` reads `D.TEMPLATES[name]` at call
time (lines 198, 259, 322), so swapping the globals makes it work against the new
workbook for free. Several `test_responsive.py` probes too — they read globals and
assert *relative* movements, which are data-agnostic.

---

## Where the dataset should live

Both shapes are defensible and this is the project's call.

| | In-repo dataset layer | Outside, in a consumer |
|---|---|---|
| **For** | One canonical implementation of the substitution; the globals problem gets solved properly by whoever owns the code; `test_rescale` and `test_responsive` integrate naturally | The recreation stays exactly one dataset with 127 tie-outs asserting published labels — the README's claim stays literally true |
| **Against** | "Everything this project builds, it builds from `ibcs_data.py`" needs rewriting; two datasets to keep in step | The globals problem gets solved by monkeypatching from outside, which is fragile and duplicates knowledge of the internals |

**If in-repo, the cleaner shape is a parameter, not a second copy.** Move the
seventeen `Template` constants and their satellite globals behind a `dataset`
object; have `_CHECKS` take the dataset it is checking rather than reading globals;
have `ibcs_excel` and `ibcs_svg` read from the passed template rather than from
`D.*`. That is a real refactor of a 4,832-line module, but it *deletes* finding 1
rather than working around it — worth it only if the project intends to support
arbitrary data long-term.

**If outside**, the mechanism is an `install()` that rebinds those globals, plus a
`READ_ME` rebind for attribution, driven from a consumer script.

---

## Attribution

**SVG side — already built.** `ibcs_svg.py:834-858` carries two footers, and the
comment above them says a chart drawn from someone's own figures "is not a
recreation of anything", so it takes `FOOTER_OWN_DATA`. The choice is made by the
layout dict's `"footer"` key. **No layout in this repo sets it** — it is a hook
waiting for exactly this work.

**Excel side — no equivalent exists.** Attribution is the hard-coded `READ_ME` list
at `ibcs_excel.py:1593`, ending:

> "Recreated for study from the IBCS® published templates. Not an IBCS Institute
> publication and not endorsed by it."

`write_read_me` does `lines = list(READ_ME)` at line 1656 — a global looked up at
call time, so it is overridable without a fork. **Consider adding a proper
`READ_ME_OWN_DATA` constant** as the Excel-side counterpart of `FOOTER_OWN_DATA`;
the asymmetry is a gap regardless of what happens to this proposal.

Keep `READ_ME` paragraphs 2-7 — the live-workbook, notation and UN 4.1 explanations
are true of any data. Replace the title, insert a "these figures are not a real
company" paragraph (or a provenance line, for real data), and replace the final
attribution with something that credits the standard while dropping the recreation
claim:

> "Drawn in IBCS® notation from [invented figures / <source>]. IBCS® is a registered
> trademark of the IBCS Association; this is not an IBCS Institute publication and is
> not endorsed by, affiliated with, or produced in cooperation with it."

`Template.source_ref` and `notes` should be reset for provenance. Verified that
nothing reads them — not `ibcs_excel`, not `ibcs_svg`, not `ibcs_doc`, not
`ibcs_guide` — so they carry no build risk. `ibcs_guide.py` and `ibcs_doc.py`
contain no Institute attribution string; nothing to change there.

---

## Per-template notes

> **Superseded.** These notes assume invented figures authored to fit the
> layouts. The shipped dataset is real and could not be authored to fit, so the
> magnitude advice below did not apply and the layout constants moved instead —
> `tier_bounds`, `panel_geometry`, `TreeSpec.scale_px`,
> `StructurePanel.number_format` and `Row.ratio_of` are now declared on the
> `Template` when a layout constant would otherwise describe the figures. C13D
> was also reframed to 15 year-panels × 12 months rather than keeping the
> fifteen location keys; the state data went to C02A and C06F.

*Source-dependent: assumes invented data. The structural observations hold
regardless.*

**General.** State the minimum and derive the rest — the typed columns are already
declared by the layouts (`Column(..., typed=True)`) and are exactly what a reader
pastes over. Reuse the existing helpers: `_derive_py`, `_variances`, `_split`,
`waterfall_spans`, `_running`, `_accumulate`, `_c11a_quotient`, `excel_round`. Set
`Tier.printed = None` throughout — it exists to reproduce the Institute's own
rounding, and against other data it is just a second place for a number to be
wrong. Build every message sentence as an f-string over derived values so it cannot
go stale.

**Tier stacks — C03A, C04A, C05X, C06F, C12A.** Two stated series, three derived
tiers. C04A's rows must be re-sorted by ΔPL descending — that ordering *is* what
makes it variant A. C06F should derive all five summary totals rather than typing
them (upstream types them because it is transcribing a picture) and assert the
bridge lands: `PY_total + ΣΔPY == AC_total`.

**Tables — T01B, T02A, T03A, T04A.** The `Row` tuple is **frozen**: same count, same
`kind`, same `spans`, same `ratio_of`, same margin/result/revenue indices. Labels
and figures change; geometry does not. T01B: restate `red_below` in the new unit —
the footnote prints it and UN 5.2 requires the threshold stated where it is applied.
T02A: upstream sets `T02A_NOV_PL = T01B_NOV_PL`; preserve that identity and check
it, so the two sheets stay one dataset seen twice. T03A: drop `T03A_SOURCE_ERROR` —
it records an Institute misprint, and there is no misprint in other data.

**Structure — C01A, C02A.** C02A: drop `C02A_SUPPRESSED` and
`C02A_PRINTED_SUBTOTALS`/`_SHARES` (label-collision facts about the reference
render, which Excel does not read). Watch `StructureLayout.max_categories` (5 for
C01A, 21 for C02A) — row and column counts are frozen.

**Line — C07C, C08H.** C07C: **MAT cannot be derived** — it reaches into the prior
year — so state it and tie it to the cumulative at December. Keep MAT the largest
series to preserve the ~4% headroom the layout assumes. C08H: state opening +
increase + decrease, derive change and levels with `_accumulate`; `D.C08H_OPENING`
and `_LEVELS` are two of the globals from finding 1.

**XY — C09C, C10D.** *Expensive.* Both need **new `Axis` declarations** — `Axis` is
data on `Template.axes` (its docstring argues exactly why: where an axis starts and
stops is an editorial decision that changes what the picture says), and
`ibcs_excel.py:3792` reads `axis.maximum` straight through. So this is data work,
not code. These two are excluded from the span-cell normalising scheme because they
are the only charts that *show* a value axis — divide on an XY chart and the reader
sees `0.25` where the data says `29.16%` — so there is no safety net here. C09C maps
onto most domains without a code change, since its derived third measure is just
`y * x / 100`. **Cutting the point count from 149 to ~70 is the single biggest
saving available** — `categories` derives from `points`, so nothing downstream
cares.

**Tree — C11A.** `C11A_SCALE_PX` is SVG-only. The workbook reads
`L.C11A_TREE_SHEET.points_per_unit`, and `ibcs_excel.py:4251` computes
`plot_height = (high - low) * per_unit` — **box height comes from the data range, so
the Excel tree auto-sizes.** No pixel refit is needed; keep the three base measures'
*ranges* near the originals and guard the fit. Separately: upstream's message
deliberately reproduces an Institute error (it names 2024 while the oval sits on
2026). A new one should name the year the oval sits on, with a tie-out asserting the
two agree.

**Panel grid — C13D.** *Expensive.* **Keep the fifteen panel keys unchanged.** Both
`PanelGrid` objects address cells by key and both are frozen, so keeping the keys
means `C13D_GRID_UNIFORM` — the one the workbook uses — works untouched. Change only
`label` and `values`: 15 × 12 = 180 values, the largest transcription in the set.
Max panel variance ≤ +360% (`PanelLayout.maximum`). Requires the companion
panel-charts skill via `os.environ.get("PANEL_CHARTS")` — note no `$` prefix in the
name.

**C12A specifically.** `C12A_REL_LIMIT = (-35.0, 265.0)` is fixed in
`ibcs_layout.py:939`, and the relative tier **clips rather than rescales, by design**
— normalising it would put the largest bar off the panel whatever the figures were,
because the largest value in any dataset plots at exactly 1.0. See `decisions.md`,
"The tiers this cannot work for, and why". The label formula
(`ibcs_layout.py:1023-1033`) is fully live, printing the true value plus
`REPT(UNICHAR(9658), MIN(3, MAX(1, INT(v/265))))`, so triangles appear and vanish as
a reader edits. Nothing needs building — data just has to be authored knowing those
bounds. Ideally two rows exceed +265% on small bases, one landing in `[265, 530)`
(one triangle) and one in `[530, 795)` (two), so the "count rises with the overrun"
behaviour is visible in the shipped file. Make `C12A_OUTLIERS` **derived** from
`L.C12A_REL_LIMIT` so the constant cannot disagree with the data.

---

## Tie-outs

*Strategy is source-independent.* Target ~85-110 checks, in the same
`check(name, condition, detail)` / `[PASS]` / raise-on-first-failure shape as the
existing ones.

**Every family.** Derived variance equals `actual − reference`; relative equals
`absolute / reference × 100` to `<1e-9`; every printed label equals `excel_round` of
its own derived value; no two scenarios claim the same category slot; every series
is the full category length.

**Waterfalls and statements** (C12A, C06F, T03A, T04A, C02A). Run `waterfall_spans`
and assert every subtotal equals the lines above it, once per scenario. Cost lines
classify adverse without being listed as such.

**Tables.** Every subtotal walks; the ratio row carries no variance; the ratio
equals numerator/denominator at `ratio_of`; each block closes independently.

**Structure.** Bands sum to the printed totals; sibling panels total the same base
period; `len(categories) <= layout.max_categories`.

**Line.** Cumulative equals the running sum, both scenarios; levels reproduce from
`opening + Σ change`; `change == increase − decrease` every period.

**XY.** Every point strictly inside its declared axis range, with headroom.

**Tree.** Every link evaluates — `ros = return/net_sales`,
`turnover = net_sales/capital`, `roi = ros × turnover` — to `<1e-9`, all periods.

**Panel grid.** Average is the mean of the fifteen; every variance reproduces; all
fifteen keys draw; variances inside the panel bounds.

### Drop these — assertions about the published picture, not arithmetic

`T03A_SOURCE_ERROR` · C09C "45 stated, 44 drawn" · C13D "12 more" vs "11 more" ·
`C13D_MEASURED_STEMS` / `_STEM_FIT` · `C11A_MEASURED_EXTENTS` / `_RETURN_RECOVERY` ·
`C10D_MEASURED_RADII` · `C02A_SUPPRESSED` / `_PRINTED_SUBTOTALS` / `_SHARES` · C06F
"Illinois derives 457 where the reference prints 456".

These are not failures of the new dataset. They are true statements about a picture
the new dataset is not a picture of.

### Three guards with no upstream counterpart

Each turns a silent visual failure into a build failure. **These matter more, not
less, if the data is real** — real figures cannot be nudged to fit.

1. **C12A clip guard.** Derived outliers non-empty, and **no subtotal row is an
   outlier** — a clipped subtotal reads as a broken statement.
2. **C11A box-fit guard.** For each node,
   `range × points_per_unit[group] / original_plot_height` within `0.6 … 1.5`,
   against the shipped `box_heights` as the baseline.
3. **XY axis-containment guard.** An out-of-range point in Excel is drawn *on the
   chart frame* rather than dropped, which looks like a rendering bug and is really
   a data bug. **Write this before authoring C09C or C10D, not after** — it is the
   difference between a one-line assertion failure and a two-hour debugging session.

---

## Verification

*Source-independent.*

| # | Check | Proves |
|---|---|---|
| V1 | `git status --short` on any repo treated as read-only — must be empty. Use `python -B` so not even a `.pyc` lands. | The dependency contract held |
| V2 | Hash the three site SVGs, re-render, compare | A refactor changed nothing already published |
| V3 | Run the new tie-out pass standalone, no Excel | The arithmetic of all seventeen. **The real gate** — `check_ties` is vacuous here |
| V4 | Build the workbook with `--doc` | Seventeen sheets, geometry verified, page setup 1×1, attribution landed. Watch for `Verification: tiers aligned.` and **no** `Verification problems:` block, which covers `check_scale()` and `check_captions()` |
| V5 | `ibcs_doc.py --xlsx <book> --check` | Every formula and address the guide quotes is what the workbook holds |
| V6 | `test_rescale.py` against the new workbook, routed through the swap so it reads the new `D.TEMPLATES` | The picture is invariant under a ×100 paste — the check other-data most needs, since the magnitudes are no longer the ones the axes were measured for |

**V7 — `test_responsive.py`: do not use as a gate.** Its C03A probe is tuned to the
published June figure (`EDIT_ROW = LAYOUT.first_row + 5`, `NEW_MEASURE = 150.0`,
asserting June starts green and turns red). Against other data June may already be
red, so the probe fails on *correct* data. Sixteen further probes each name a
specific row and delta. Most read module globals and assert relative movements, so
the swap makes them data-agnostic for free — the colour-flip one does not. Either
retune it or drop it; its value here is redundant with V6 plus the manual pass, and
it is this project's test of this project's workbook.

**V8 — manual Excel pass.** Read me carries the new attribution and no "Recreated
for study" · `Ctrl+F` for `Furniture`, `Alpha`, `Beta`, `Berlin`, `Cologne`,
`California`, `kEUR`, `kCHF`, `mUSD` returns zero hits · C12A outliers show `►` /
`►►`, stop at the panel edge, and lose the triangles when the value is edited back
inside · C11A six boxes at sensible heights on one page, oval on the year the
message names · C13D 4×4 grid with one empty cell and the outlier panel contained ·
C09C/C10D every point inside the frame, nothing drawn on the border · retype one
input per family and confirm the chart moves and a sign change changes colour ·
print preview one page on three sheets.

---

## Effort

> **Superseded.** An estimate for authoring invented figures. The shipped work
> harvested real ones from SEC EDGAR instead (`fetch_pgr.py`, `pgr_parse.py`),
> which moved the cost from typing numbers to parsing filings and reconciling
> them. Do not read the table below as a record of what this took.

*Source-dependent: assumes invented data.* ≈ **34 hours**, dominated by data
authoring rather than code.

| | |
|---|---|
| Harness, globals swap, attribution, three-template proof | 3.5 h |
| Cheap — C04A, C05X, C06F, T01B, T02A, T03A, C02A | 8.5 h |
| Medium — C12A, C07C, C08H, C11A | 8 h |
| Expensive — C09C (4 h), C10D (3 h), C13D (5 h) | 12 h |
| Guide, packaging, verification and iteration | 4 h |

**Cheapest:** the tier stacks — two stated series, everything derived, structure
frozen, and `render-site-svgs.py`'s C03A is a working example to copy.

**Most expensive:** C13D (180 values, five globals, the only external repository
dependency) · C09C (~70 points, six globals, two new axes, iso-curve arithmetic that
must stay consistent with `c09c_gross_profit`) · C10D (three simultaneous measures
with a bubble-area law over two freshly declared axes, with no span-cell safety
net).

**Sequence that keeps each step verifiable:** harness first with the three existing
templates only, and gate it on the site's three SVGs coming out byte-identical — a
free regression test of the whole extraction before a single new number is typed.
Then the attribution and build path, still on three templates. Then cheap families,
C12A, the globals-patching families, and the two expensive ones last, C13D last of
all because of its external dependency.

---

## Open question — decided: no

> **Decided 2026-08-21: the website keeps its own images.**
> `render-site-svgs.py` continues to carry its own substituted figures for
> C03A, C01A and T04A, and is not repointed at `ibcs_data_alt`.
>
> The website kept its original images, with an indicator that they contain
> synthetic data. The decision was made because the charts were visually
> attractive.
>
> The indicator is already in place three ways: the entity is named *Sample
> Mutual Insurance*, each image carries a `· synthetic figures` caption on the
> page, and the SVGs take `FOOTER_OWN_DATA` rather than the recreation footer.
> So the drift this section warns about cannot mislead a reader — nobody can
> mistake these three pictures for the named insurer's real filings in the
> workbook beside them, which was the actual risk.
>
> `render-site-svgs.py` carries the same decision in its own header, with the
> instruction not to "fix" it by importing the dataset. The paragraph below is
> the argument that was made for repointing it; it did not win.

`scripts/render-site-svgs.py` in the `edgewisedata` website repo currently owns
substituted data for three templates and generates the images that site publishes.
If this project takes on a dataset, **that script should probably consume it rather
than carry its own** — otherwise the site's images and this project's workbook are
two independent sets of invented numbers that can drift, and the site ends up
showing pictures of figures its download does not contain.

Small change if decided up front. Annoying one later.
