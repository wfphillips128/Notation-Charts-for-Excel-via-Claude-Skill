# Drawing the templates from figures that are not the Institute's

The recreation is a fidelity exercise: every figure in `ibcs_data.py` is
transcribed from the IBCS® Institute's published example renderings, and that
is the point of it. It makes a poor file to paste your own numbers into,
because you cannot tell which figures are Furniture Inc.'s without checking
each one.

So the seventeen templates are drawn a second time, from **The Progressive
Corporation (PGR)** - a real insurer's real, public figures. This is how that
works, and what to copy if you want a third.

## What ships

Everything needed to build the Progressive workbook is committed. **A build
never touches the network.**

| | |
|---|---|
| `datasets/pgr/*.csv` | five files, the figures |
| `scripts/ibcs_data_alt.py` | the seventeen `Template` objects, built by reading those CSVs |
| `scripts/build_alt.py` | the entry point |
| `scripts/fetch_pgr.py`, `scripts/pgr_parse.py` | how the CSVs got there. Not needed to build |
| `scripts/test_alt_ties.py`, `scripts/test_pgr_reconcile.py` | the two gates this dataset adds |

## The entry point

```bash
python scripts/build_alt.py pgr.xlsx --doc "pgr - Excel Guide.md"
python scripts/build_alt.py pgr-simple.xlsx --simple
python scripts/build_alt.py out.xlsx --template C03A,C09C --export-dir pics/
```

| Argument | |
|---|---|
| `out` | positional, default `IBCS-Progressive.xlsx` |
| `--simple` | the base-tier form, for the templates that have a useful one |
| `--template` | comma-separated ids; default is every template the dataset carries |
| `--doc` | write the by-hand build guide beside the workbook |
| `--export-dir` | a PNG per sheet's charts |

**`build_alt.py` exists because `ibcs_excel.py`'s CLI cannot be pointed
elsewhere.** That CLI resolves template ids against `ibcs_data`. `build()`
itself has always taken the templates it is handed - so `build_alt` swaps
nothing, monkeypatches nothing, and imports no `sys.modules` trickery. That is
the whole property, and it is what `test_dataset_isolation.py` protects.

**Reach for `--export-dir` when the panel grid is in the build.** Nothing in
either repository measures a rendered label's bounding box, so a value label
printing through a rule passes every gate - the digits are struck through, not
missing. Exporting the picture and looking at it is the check.

**A template the dataset does not carry is an error, not an empty sheet.**
`build_alt` prints what the dataset *does* carry rather than raising `KeyError`,
and `--simple` refuses ids that have no simple form. Which templates have one is
computed from `SIMPLE_LAYOUTS` and the structure charts rather than retyped -
two ways of knowing the same fact is one of them going stale.

## Harvesting, which you do not need to do

**The CSVs are committed. `fetch_pgr.py` is only how they got there.** Run it to
extend the range or to re-derive after a parser change; never as a build step.

**Set `SEC_USER_AGENT` first.** SEC's fair-access policy requires every request
to name a real person and a contact address, and *refuses* requests without one
- HTTP 403, not a throttle. The refusal reads like a network fault, which is how
it costs an afternoon.

```powershell
$env:SEC_USER_AGENT = 'Your Name you@example.com'
```

It is read from the environment and never hard-coded, because this repository is
public.

```bash
python scripts/fetch_pgr.py --probe                    # what EDGAR holds; index only
python scripts/fetch_pgr.py --harvest                  # pull releases, write the CSVs
python scripts/fetch_pgr.py --harvest --since 2015-01-01 --until 2025-12-31
```

`--since` and `--until` default to `2011-01-01` and `2026-12-31` and carry no
help text of their own.

Two more facts about EDGAR that are not obvious: **ten requests a second is the
ceiling**, so everything goes through one throttled door at 0.15 s; and
**nothing is fetched twice** - responses land in `build/pgr-cache/`, which the
repo ignores, so re-parsing costs no requests.

## Parsing

`pgr_parse.py` turns one monthly news release into figures. Two things about it
are worth copying.

**It keys on row labels, not positions.** Fifteen years of releases are not
identically formatted. Positions drift; labels do not.

**The prior-year column is why this is parsed rather than typed.** Every release
prints the same month a year earlier beside the current one, so every figure is
checkable against the release that first reported it twelve months back. A
transcription error has to be made twice, identically, a year apart, to survive.

## The CSV contract

| File | Rows | Holds |
|---|---|---|
| `monthly_headline.csv` | 211 | companywide monthly NPW, NPE, net income, combined ratio, each with its prior-year column, plus the accession number and filing date |
| `monthly_segments.csv` | 1,044 | the same months split by segment: Agency, Direct, Property, Personal Lines, Commercial, Companywide |
| `ytd_segments.csv` | 978 | year-to-date by segment; the `through_month == 12` rows are the only per-segment annual figures anywhere in the archive |
| `states.csv` | 239 | NPW by state, from the 10-K state tables |
| `xbrl_facts.csv` | 2,085 | the fourteen XBRL concepts four templates read - C08H, C11A, C12A and T03A - as `concept,start,end,val` in filing order |

**A missing figure is an empty cell, never the string `None`.**

**`xbrl_facts.csv` is in filing order and that is load-bearing.** Later filings
restate earlier years; the readers build dicts and let the last one win. Re-sort
it and the restatements stop applying.

It exists because `ibcs_data_alt` used to read the 5 MB companyfacts JSON out of
the harvest cache, and `build/` is gitignored - so the dataset could not be
built from a clean clone at all. Fourteen concepts are ever read from that file.
They ship; the 5 MB does not.

## The dataset module

`ibcs_data_alt.py` reads the CSVs and exports `TEMPLATES`, seventeen ids, same
keys as `ibcs_data`. It transcribes nothing by hand.

**Eleven of the seventeen templates want a plan or a forecast, and no public
company files one.** Those are constructed from Progressive's own filed target -
`TARGET_COMBINED_RATIO = 96.0`, so a 4% underwriting margin, so plan
underwriting profit is `net premiums earned x 4 / 100`. The rule is printed on
each sheet that uses it, and every such scenario is declared `assumed`. See
`reference-provenance.md`.

## The four gates

| Gate | Excel? | Proves |
|---|---|---|
| `test_alt_ties.py` | no | 211 tie-outs, each reading the template it was handed |
| `test_pgr_reconcile.py` | no | 605 checks against Progressive's own filed totals |
| `test_dataset_isolation.py` | no | no renderer reaches past its template |
| `test_provenance.py` | **yes** | constructed figures are shaded, and nothing else is |

**`test_alt_ties` is the real gate, because the usual one is not.**
`ibcs_data.check_ties()` takes a `Template` and then ignores it - every
per-template check reads module globals from `ibcs_data`. Handed a template
built from other figures it re-verifies the Institute's numbers against
themselves, passes, and proves nothing about the data being drawn. `build()`
gates on it, so without a replacement the workbook would be built on an
assertion that cannot fail.

`build()` takes a `check_ties` parameter for exactly this, and `build_alt`
passes `test_alt_ties.check_template`.

**The gate must refuse what it does not know.** `check_template` raises
`no tie-outs written for {name}; refusing to build it` on any template it has no
checks for. That is what stops a sheet being added without its arithmetic.

**`test_pgr_reconcile` makes the claim invented figures never could.** Three
independent tests, each of which a transcription error has to survive: twelve
monthly premium figures from twelve separate news releases summing to the annual
number tagged in the audited 10-K; every release's prior-year column agreeing
with the release of twelve months before; segments adding to companywide. A
reconciliation *against the source* is a stronger claim than internal
consistency, and it is what real data buys that authored data cannot.

A tolerance is allowed on the annual sum only, because twelve figures each
rounded to a million can legitimately differ from one figure rounded once. It is
tight enough that a wrong month cannot hide in it.

**`test_dataset_isolation` runs on the AST and needs nothing installed.** Every
`ibcs_data` name a renderer reads must be a type, a helper, a registry or a
`Template`. Anything else is data, and data must arrive through the argument.
Written against the live module rather than a list of forbidden names, so a new
global fails the day it is written.

## Doing this for your own figures

1. **Write a reader**, not a transcription. A module that reads committed files
   and builds `Template` objects. If the figures came from somewhere, the thing
   that fetched them ships too, and the fetch is never a build step.
2. **Declare `provenance` on every template** that carries a scenario nobody
   published. `reference-provenance.md`.
3. **Declare `tier_bounds` and `panel_geometry` as fractions of the span**, on
   the `Template`. A layout constant measured against one company's magnitudes
   is wrong for yours - see the layout-constant rule in `SKILL.md`.
4. **Write per-template tie-outs** exposing `check_template(t)`, refusing
   templates it does not cover.
5. **Pass it in**: `build(templates, out, check_ties=your_gate)`.
6. **Reconcile against the source** if the figures are real. It is the strongest
   check available and it is only available to you.
7. **Two disclaimers on the Read me**, and export the pictures and look at them.
