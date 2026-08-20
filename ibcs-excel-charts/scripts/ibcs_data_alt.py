"""The same seventeen templates, drawn from The Progressive Corporation.

``ibcs_data`` is a fidelity recreation: every figure in it is transcribed from
the IBCS(R) Institute's published example renderings, and that is the point of
it. It makes a poor file to paste your own numbers into, because you cannot
tell which of them are Furniture Inc.'s without checking each one.

This module is the other thing people keep wanting - the same templates, the
same renderer, on figures that came from somewhere real and say where. Every
number here is either read from a Progressive filing or derived from ones that
were, and each template declares which through ``Template.provenance``, so the
workbook can shade a constructed figure differently from a reported one.

Nothing is transcribed by hand. ``fetch_pgr.py`` pulls the filings, and this
reads the CSVs it writes - which is why ``test_pgr_reconcile.py`` can assert
that twelve monthly figures sum to the audited annual before any of this runs.

**Not affiliated with, endorsed by, or produced in cooperation with The
Progressive Corporation.** Figures are from public SEC filings and
investor-relations releases; plan and forecast scenarios are constructed here on
stated bases and were never published by the company.
"""
from __future__ import annotations

import csv
from pathlib import Path

import ibcs_data as D
import ibcs_layout as L

DATA = Path(__file__).resolve().parents[1] / "datasets" / "pgr"

ENTITY = "The Progressive Corporation"

# Progressive's long-stated underwriting target: a 96 combined ratio, which is
# a 4% underwriting margin. It is *filed* - stated in every 10-K and the basis
# of the company's Gainsharing programme. Applying it to a period to produce a
# plan figure is this workbook's construction and is marked ``assumed``.
TARGET_COMBINED_RATIO = 96.0
TARGET_MARGIN = 100.0 - TARGET_COMBINED_RATIO


def _rows(name: str) -> list[dict]:
    with (DATA / name).open(encoding="utf-8") as fh:
        return list(csv.DictReader(fh))


def _num(value: str | None) -> float | None:
    if value in ("", None, "None"):
        return None
    return float(value)


def monthly() -> dict[tuple[int, int], dict]:
    """Companywide monthly figures, keyed by (year, month)."""
    return {(int(r["year"]), int(r["month"])): r
            for r in _rows("monthly_headline.csv")}


def annual_segments() -> dict[int, dict[str, dict]]:
    """Full-year segment figures, keyed by year then role.

    December's year-to-date block is the whole year, and it is the only place
    annual figures exist *per segment* anywhere in the archive - the 10-K tags
    premium companywide, not by segment.
    """
    out: dict[int, dict[str, dict]] = {}
    for r in _rows("ytd_segments.csv"):
        if int(r["through_month"]) == 12 and r["role"]:
            out.setdefault(int(r["year"]), {})[r["role"]] = r
    return out


# --------------------------------------------------------------------------- #
# C03A - monthly premium against prior year, with a forecast tail
# --------------------------------------------------------------------------- #

C03A_YEAR = 2026
C03A_PRIOR = 2025

_M = monthly()
C03A_PY_VALUES = tuple(_num(_M[(C03A_PRIOR, m)]["npw"]) for m in range(1, 13))
_actual = [(_num(_M[(C03A_YEAR, m)]["npw"]) if (C03A_YEAR, m) in _M else None)
           for m in range(1, 13)]
C03A_LAST_ACTUAL = max(i for i, v in enumerate(_actual) if v is not None) + 1

# The forecast rule, stated here and printed on the sheet. Premium is not a
# ratio, so the 96 combined ratio target says nothing about it; the honest
# construction is that the rest of the year grows on the same year-to-date pace
# the reported months already show.
_ytd_ac = sum(v for v in _actual[:C03A_LAST_ACTUAL] if v is not None)
_ytd_py = sum(C03A_PY_VALUES[:C03A_LAST_ACTUAL])
C03A_PACE = _ytd_ac / _ytd_py

C03A_MEASURE = tuple(
    v if v is not None else D.excel_round(C03A_PY_VALUES[i] * C03A_PACE)
    for i, v in enumerate(_actual))
C03A_SCENARIOS = tuple("AC" if i < C03A_LAST_ACTUAL else "FC" for i in range(12))

C03A_VAR_ABS = tuple(a - p for a, p in zip(C03A_MEASURE, C03A_PY_VALUES))
C03A_VAR_REL = tuple(v / p * 100.0 for v, p in zip(C03A_VAR_ABS, C03A_PY_VALUES))

_total_ac = sum(v for v, s in zip(C03A_MEASURE, C03A_SCENARIOS) if s == "AC")
_total_fc = sum(v for v, s in zip(C03A_MEASURE, C03A_SCENARIOS) if s == "FC")
_total_py = sum(C03A_PY_VALUES)
_total_var = (_total_ac + _total_fc) - _total_py


def _split(values, keep):
    return [v if s == keep else None for v, s in zip(values, C03A_SCENARIOS)]


C03A = D.Template(
    id="C03",
    variant="A",
    kind="multi-tier column chart",
    title=D.TitleBlock(
        entity=ENTITY,
        measure="Net premiums written",
        unit="$m",
        period=str(C03A_YEAR),
        # Built from the derived figures rather than typed, so it cannot go
        # stale when the underlying months are re-harvested.
        message=(
            f"Net premiums written are running "
            f"{min(C03A_VAR_REL[:C03A_LAST_ACTUAL]):+.1f}% to "
            f"{max(C03A_VAR_REL[:C03A_LAST_ACTUAL]):+.1f}% above prior year "
            f"through {D.MONTHS[C03A_LAST_ACTUAL - 1]}; on that pace the year "
            f"closes {_total_var:+,.0f} $m ({_total_var / _total_py * 100:+.1f}%) "
            f"above {C03A_PRIOR}"
        ),
    ),
    categories=D.MONTHS,
    category_scenarios=C03A_SCENARIOS,
    tiers=(
        D.Tier(
            key="var_rel", label="ΔPY%", kind="variance_rel", reference="PY",
            height_weight=0.85, number_format="{:+.1f}",
            series=(D.Series("AC", _split(C03A_VAR_REL, "AC")),
                    D.Series("FC", _split(C03A_VAR_REL, "FC"))),
        ),
        D.Tier(
            key="var_abs", label="ΔPY", kind="variance_abs", reference="PY",
            height_weight=0.85, number_format="{:+,.0f}",
            # Deliberately None. ``printed`` exists to reproduce the Institute's
            # own rounding; on other data it is just a second place for a number
            # to be wrong.
            printed=None,
            series=(D.Series("AC", _split(C03A_VAR_ABS, "AC")),
                    D.Series("FC", _split(C03A_VAR_ABS, "FC"))),
        ),
        D.Tier(
            key="measure", label="", kind="measure", height_weight=2.6,
            number_format="{:,.0f}",
            series=(D.Series("PY", C03A_PY_VALUES, label="PY"),
                    D.Series("AC", _split(C03A_MEASURE, "AC"), label="AC"),
                    D.Series("FC", _split(C03A_MEASURE, "FC"), label="FC")),
        ),
    ),
    summary_rows=(D.Summary(
        label=str(C03A_YEAR),
        stack=(("AC", D.excel_round(_total_ac)), ("FC", D.excel_round(_total_fc))),
        variance_abs=D.excel_round(_total_var),
        variance_rel=D.excel_round(_total_var / _total_py * 100.0, 1),
        reference="PY",
    ),),
    # Axis bounds as fractions of each tier's span, which is the form the scale
    # block actually plots in. The recreation declares its bounds in kEUR and
    # divides them by the live span; those two only agree when the span is the
    # one the bounds were measured against. Progressive's premiums are ~47x the
    # Institute's contribution figures, so that route put the measure axis at
    # 0.023 while every normalised bar sat near 1.0 - Excel clamped them all to
    # full height and the chart read as flat while the data varied fourfold.
    #
    # These are the recreation's own bounds divided by its own span, so the
    # picture keeps the proportions the reference render has. Being fractions,
    # they hold for any figures at all - which is what a workbook meant to be
    # pasted over needs.
    tier_bounds={
        "measure": (0.0, 230.0 / 211.0),        # 0 .. 1.090
        "var_abs": (-36.5 / 211.0, 55.5 / 211.0),   # -0.173 .. 0.263
        "var_rel": (-15.0 / 17.3, 23.0 / 17.3),     # -0.867 .. 1.329
    },
    provenance=(
        D.SourceNote("AC", "filed",
                     "Monthly net premiums written, Exhibit 99 to the monthly "
                     "8-K. Twelve months reconcile to the 10-K annual figure."),
        D.SourceNote("PY", "filed",
                     f"The same series for {C03A_PRIOR}."),
        D.SourceNote("FC", "assumed",
                     f"Progressive publishes no forecast. Remaining months are "
                     f"prior-year months grown at the year-to-date pace through "
                     f"{D.MONTHS[C03A_LAST_ACTUAL - 1]} "
                     f"({C03A_PACE * 100 - 100:+.1f}%)."),
    ),
    source_ref="SEC EDGAR CIK 0000080661, monthly 8-K exhibits",
    notes=(
        "Two reported series - this year and last - and everything else "
        "derived from them.",
        "The forecast tail is constructed, not reported, and is shaded and "
        "footnoted as such.",
    ),
)


# --------------------------------------------------------------------------- #
# C10D - underwriting performance by segment, as a bubble portfolio
# --------------------------------------------------------------------------- #

C10D_AC_YEAR = 2016
C10D_PY_YEAR = 2015

_SEG = annual_segments()
C10D_LABELS = {"agency": "Agency auto", "direct": "Direct auto",
               "commercial": "Commercial", "property": "Property"}

# Property is drawn as an acquisition rather than as a prior-year pair, and the
# reason is in the harvest rather than in a press release: its first monthly row
# anywhere in the archive is April 2015, when Progressive took majority control
# of ARX. Its 2015 figure therefore covers nine months, not twelve. The ratios
# would survive that - a loss ratio does not care how long the period is - but
# the bubble *area* is earned premium, and a nine-month bubble beside twelve-
# month ones is a comparison that means nothing. A unit with no comparable prior
# year, drawn without a prior-year twin, is exactly what the ACQ scenario says.
C10D_ACQUIRED = "property"


def _c10d_points() -> tuple[D.Point, ...]:
    out: list[D.Point] = []
    for role, label in C10D_LABELS.items():
        for year, scenario in ((C10D_PY_YEAR, "PY"), (C10D_AC_YEAR, "AC")):
            if role == C10D_ACQUIRED and scenario == "PY":
                continue
            row = _SEG.get(year, {}).get(role)
            if not row or not row["combined_ratio"]:
                continue
            out.append(D.Point(
                entity=label,
                scenario="ACQ" if role == C10D_ACQUIRED else scenario,
                x=float(row["expense_ratio"]),
                y=float(row["loss_ratio"]),
                size=float(row["npe"]),
            ))
    return tuple(out)


C10D_POINTS = _c10d_points()
_ac = {p.entity: p for p in C10D_POINTS if p.scenario in ("AC", "ACQ")}
_worst = max(_ac.values(), key=lambda p: p.x + p.y)
_best = min(_ac.values(), key=lambda p: p.x + p.y)
# The interesting fact, and it has to be *checked* rather than asserted: the
# segment with the worst combined ratio may well have the best loss ratio, in
# which case the whole story is on the expense axis.
_lowest_loss = min(_ac.values(), key=lambda p: p.y)
_worst_is_leanest_on_losses = _lowest_loss.entity == _worst.entity
_expense_gap = _worst.x - min(p.x for p in _ac.values())

C10D = D.Template(
    id="C10",
    variant="D",
    kind="bubble chart",
    title=D.TitleBlock(
        entity=ENTITY,
        measure="Underwriting performance by segment",
        # Three measures in three units, so none of them can head the chart;
        # each is named where it is used.
        unit="",
        period=str(C10D_AC_YEAR),
        message=(
            f"{_worst.entity} carries the highest combined ratio at "
            f"{_worst.x + _worst.y:.1f}"
            + (f" despite the lowest loss ratio in the group at {_worst.y:.1f}"
               if _worst_is_leanest_on_losses else
               f", against {_best.entity} at {_best.x + _best.y:.1f}")
            + f"; its expense ratio of {_worst.x:.1f} is "
              f"{_expense_gap:.0f} points above the leanest segment"
        ),
    ),
    categories=tuple(p.entity for p in C10D_POINTS),
    category_scenarios=("AC",) * len(C10D_POINTS),
    tiers=(),
    points=C10D_POINTS,
    # Bounds are nominal: an XY template is fitted to its own points at build,
    # keeping the step the template chose. The step is the real decision here -
    # how finely a reader should be able to read a ratio.
    # Units kept short. The axis captions are written into the data zone as
    # cells, and "% of earned premium" twice ran the second one past the chart
    # boundary - the caption is the one piece of text on an XY sheet long
    # enough to do that. Both ratios are of earned premium by definition, and
    # the notes below say so.
    axes=(D.Axis("Expense ratio", 15.0, 35.0, 5.0, "{:.0f}", unit="%"),
          D.Axis("Loss and LAE ratio", 55.0, 80.0, 5.0, "{:.0f}", unit="%")),
    size_legend=("Net premiums earned", "$m"),
    # The two auto segments sit almost on top of each other in both years, and
    # four labels in that corner are unreadable. The prior-year names go: the
    # light bubble still shows where the segment was, and the actual beside it
    # is named. Every value label is kept - it is only the names that give way,
    # which is the rule the reference follows too.
    unlabelled=frozenset({("Agency auto", "PY"), ("Direct auto", "PY")}),
    provenance=(
        D.SourceNote("AC", "filed",
                     f"Full-year segment results for {C10D_AC_YEAR}, from the "
                     f"year-to-date block of the December monthly release."),
        D.SourceNote("PY", "filed",
                     f"The same block for {C10D_PY_YEAR}."),
        D.SourceNote("ACQ", "filed",
                     f"Property's {C10D_PY_YEAR} figures cover nine months - "
                     f"ARX was acquired in April {C10D_PY_YEAR} - so it is "
                     f"drawn without a prior-year twin rather than against a "
                     f"period it cannot be compared with."),
    ),
    source_ref="SEC EDGAR CIK 0000080661, monthly 8-K exhibits",
    notes=(
        "Both axes are components of the combined ratio, so a diagonal is a "
        "line of constant underwriting result; Progressive's stated target is "
        "a 96 combined ratio.",
        "Bubble area is earned premium, never radius - the eye reads area.",
    ),
)


# --------------------------------------------------------------------------- #
# C02A - where the premium dollar goes, by segment
# --------------------------------------------------------------------------- #

C02A_YEAR = 2025

# A structure chart needs a breakdown that *adds up*, and Progressive publishes
# no state-by-segment cut - the state table is premium only. What it does give,
# per segment, is the two ratios that consume earned premium: everything left
# after losses and expenses is the underwriting result. Those three are the
# premium dollar, they sum to it exactly, and they aggregate up the segment
# hierarchy the way a structure chart's subtotals require.
C02A_LEAVES = ("agency", "direct", "property", "commercial")
C02A_LABELS = {"agency": "Agency auto", "direct": "Direct auto",
               "property": "Property", "commercial": "Commercial"}
C02A_PERSONAL = ("agency", "direct", "property")


def _c02a_parts(role: str) -> tuple[float, float, float]:
    """(losses, expenses, underwriting result) in $m for one segment."""
    row = _SEG[C02A_YEAR][role]
    npe = float(row["npe"])
    losses = npe * float(row["loss_ratio"]) / 100.0
    expenses = npe * float(row["expense_ratio"]) / 100.0
    return losses, expenses, npe - losses - expenses


_parts = {role: _c02a_parts(role) for role in C02A_LEAVES}
_personal = tuple(sum(_parts[r][i] for r in C02A_PERSONAL) for i in range(3))
_company = tuple(sum(_parts[r][i] for r in C02A_LEAVES) for i in range(3))

C02A_CATEGORIES = ("Agency auto", "Direct auto", "Property", "Personal Lines",
                   "Commercial", "Companywide")
_C02A_COLUMNS = ([_parts[r] for r in C02A_PERSONAL] + [_personal]
                 + [_parts["commercial"], _company])
C02A_SEGMENT_LABELS = ("Losses and LAE", "Expenses", "Underwriting result")

_result = _company[2]
_worst = min(C02A_LEAVES, key=lambda r: _parts[r][2] / sum(_parts[r]))
_best = max(C02A_LEAVES, key=lambda r: _parts[r][2] / sum(_parts[r]))

C02A = D.Template(
    id="C02",
    variant="A",
    kind="stacked bar chart",
    orientation="horizontal",
    title=D.TitleBlock(
        entity=ENTITY,
        measure="Net premiums earned",
        unit="$m",
        period=str(C02A_YEAR),
        message=(
            f"Of {sum(_company):,.0f} $m earned in {C02A_YEAR}, "
            f"{_company[0]:,.0f} went to losses and {_company[1]:,.0f} to "
            f"expenses, leaving an underwriting result of {_result:,.0f} $m "
            f"({_result / sum(_company) * 100:.1f}%); "
            f"{C02A_LABELS[_best]} keeps the largest share of its premium at "
            f"{_parts[_best][2] / sum(_parts[_best]) * 100:.1f}% and "
            f"{C02A_LABELS[_worst]} the smallest at "
            f"{_parts[_worst][2] / sum(_parts[_worst]) * 100:.1f}%"
        ),
    ),
    categories=C02A_CATEGORIES,
    category_scenarios=("AC",) * len(C02A_CATEGORIES),
    # Which categories are subtotals, and it is not decoration: a horizontal
    # structure chart plots the elements and treats the subtotals as its
    # integrated legend. A bar for Companywide would be seven times the longest
    # segment and would take the shared scale with it, flattening everything
    # else - the same failure the axis guard exists for, arriving from the
    # data side.
    # ``spans`` says how many element categories a subtotal reaches back over.
    # Personal Lines takes the three before it; Companywide takes "zero", which
    # means every element so far - skipping the subtotal in between, so nothing
    # is counted twice. The workbook builds these as live SUMs, so a corrected
    # segment moves its subtotal and the grand total together.
    rows=tuple(D.Row(label=name, sign=1,
                     kind="subtotal" if name in ("Personal Lines",
                                                 "Companywide") else "element",
                     spans=(3 if name == "Personal Lines"
                            else "zero" if name == "Companywide" else None))
               for name in C02A_CATEGORIES),
    tiers=(),
    structure_panels=(D.StructurePanel(
        key="premium",
        label="",
        categories=C02A_CATEGORIES,
        category_scenarios=("AC",) * len(C02A_CATEGORIES),
        segments=tuple(
            D.Segment(label=label,
                      values=tuple(D.excel_round(col[i]) for col in _C02A_COLUMNS))
            for i, label in enumerate(C02A_SEGMENT_LABELS)),
        legend_side="left",
        legend_at=0,
        # Deliberately absent. ``printed_totals`` reproduces the figures the
        # Institute's own render prints beside each bar; here the totals are
        # derived from the segments and printing a second copy would be a second
        # place for a number to be wrong.
        printed_totals=None,
    ),),
    provenance=(
        D.SourceNote("AC", "filed",
                     f"Full-year {C02A_YEAR} segment earned premium and the "
                     f"loss and expense ratios, from the year-to-date block of "
                     f"the December monthly release. Losses and expenses are "
                     f"those ratios applied to earned premium; the underwriting "
                     f"result is what remains."),
    ),
    source_ref="SEC EDGAR CIK 0000080661, monthly 8-K exhibits",
    notes=(
        "Losses, expenses and the underwriting result are the earned premium "
        "dollar; they sum to it by construction and up the hierarchy by "
        "arithmetic.",
        "Personal Lines is Agency plus Direct plus Property; Companywide adds "
        "Commercial.",
    ),
)


# --------------------------------------------------------------------------- #
# C01A - premium by segment over five years, and by state
# --------------------------------------------------------------------------- #

C01A_YEARS = tuple(str(y) for y in range(2021, 2026))
C01A_SEGMENTS = ("agency", "direct", "property", "commercial")
C01A_LABELS = {"agency": "Agency auto", "direct": "Direct auto",
               "property": "Property", "commercial": "Commercial"}

# **Two panels, not three.** The reference cuts its measure three ways - by
# business area, by industry, by region. Progressive publishes two: by segment
# and by state. There is no customer-industry disclosure to stand in for the
# third, and inventing one would be the only figure in this workbook that came
# from nowhere. A panel fewer is the honest shape.
def _c01a_residual(year: int) -> float:
    """Companywide premium less the segments Progressive names.

    It is not always nothing. In 2021 the named columns come to 46,400.9
    against a companywide of 46,405.2 - a 4.3 residual, at one-decimal
    precision, so not rounding. Progressive's releases footnote small run-off
    businesses that no column carries, and this is them.

    Shown rather than absorbed. A structure chart's whole claim is that the
    parts make the whole; quietly dropping a remainder into the largest bar, or
    letting the stack fall short of the total printed beside it, breaks that for
    a saving of one segment.
    """
    named = sum(float(_SEG[year][role]["npw"]) for role in C01A_SEGMENTS)
    return float(_SEG[year]["companywide"]["npw"]) - named


C01A_RESIDUALS = tuple(_c01a_residual(int(y)) for y in C01A_YEARS)

C01A_PANEL_SEGMENT = D.StructurePanel(
    key="segment",
    label="By segment",
    categories=C01A_YEARS,
    category_scenarios=("AC",) * len(C01A_YEARS),
    segments=tuple(
        D.Segment(label=C01A_LABELS[role],
                  values=tuple(D.excel_round(float(_SEG[int(y)][role]["npw"]))
                               for y in C01A_YEARS))
        for role in C01A_SEGMENTS)
    + ((D.Segment(label="Other",
                  values=tuple(D.excel_round(v) for v in C01A_RESIDUALS)),)
       if any(abs(v) >= 0.5 for v in C01A_RESIDUALS) else ()),
    legend_side="left",
    legend_at=0,
    printed_totals=None,
    # The layout's "0.0" is right for figures around a hundred and prints
    # "17258.0" for figures around twenty thousand.
    number_format="# ##0",
)


def _c01a_states(year: int, keep: int = 5) -> list[tuple[str, float]]:
    """The largest states that year, with everything else gathered behind them.

    Ranked within the year rather than fixed across years: the disclosed ten
    are not the same ten each time - Louisiana was among them until the
    mid-2010s and Arizona is now - so a list fixed to one year leaves holes in
    the others. Here only one year is drawn, but the rule is the one that
    generalises.
    """
    rows = [(name, value) for name, value in _STATES[year].items()
            if name not in ("total", "all other")]
    rows.sort(key=lambda kv: -kv[1])
    rest = (_STATES[year]["total"] - sum(v for _n, v in rows[:keep]))
    return [(n.title(), v) for n, v in rows[:keep]] + [("All other", rest)]


_STATES = {}
for _r in _rows("states.csv"):
    _STATES.setdefault(int(_r["year"]), {})[_r["state"]] = float(_r["npw"])

C01A_STATE_YEAR = max(_STATES)
_state_rows = _c01a_states(C01A_STATE_YEAR)

C01A_PANEL_STATE = D.StructurePanel(
    key="state",
    label="By state",
    categories=(str(C01A_STATE_YEAR),),
    category_scenarios=("AC",),
    segments=tuple(D.Segment(label=name, values=(D.excel_round(value),))
                   for name, value in _state_rows),
    legend_side="right",
    legend_at=0,
    printed_totals=None,
    number_format="# ##0",
)

_first = {r: float(_SEG[int(C01A_YEARS[0])][r]["npw"]) for r in C01A_SEGMENTS}
_last = {r: float(_SEG[int(C01A_YEARS[-1])][r]["npw"]) for r in C01A_SEGMENTS}
_fastest = max(C01A_SEGMENTS, key=lambda r: _last[r] / _first[r])
_share = _last[_fastest] / sum(_last.values()) * 100.0

C01A = D.Template(
    id="C01",
    variant="A",
    kind="stacked column chart",
    title=D.TitleBlock(
        entity=ENTITY,
        measure="Net premiums written",
        unit="$m",
        period=f"{C01A_YEARS[0]}..{C01A_YEARS[-1]}",
        message=(
            f"{C01A_LABELS[_fastest]} has grown from {_first[_fastest]:,.0f} "
            f"to {_last[_fastest]:,.0f} $m since {C01A_YEARS[0]}, "
            f"{_last[_fastest] / _first[_fastest] * 100 - 100:+.0f}%, and now "
            f"writes {_share:.0f}% of premium; Texas and Florida together "
            f"account for "
            f"{sum(v for n, v in _state_rows[:2]) / _STATES[C01A_STATE_YEAR]['total'] * 100:.0f}%"
            f" of the {C01A_STATE_YEAR} book"
        ),
    ),
    categories=C01A_YEARS,
    category_scenarios=("AC",) * len(C01A_YEARS),
    tiers=(),
    structure_panels=(C01A_PANEL_SEGMENT, C01A_PANEL_STATE),
    provenance=(
        D.SourceNote("AC", "filed",
                     "Segment premium from the year-to-date block of each "
                     "December monthly release; state premium from the 10-K "
                     "annual-report exhibit. Both reconcile to the same "
                     "companywide total."),
    ),
    source_ref="SEC EDGAR CIK 0000080661, monthly 8-K exhibits and 10-K",
    notes=(
        "Two panels rather than three: Progressive publishes a segment cut and "
        "a state cut, and no third.",
        "The two panels share one scale, so a bar in either can be read "
        "against a bar in the other.",
    ),
)


# --------------------------------------------------------------------------- #
# C12A - the profit and loss statement as a waterfall
# --------------------------------------------------------------------------- #

C12A_AC_YEAR, C12A_PY_YEAR = 2023, 2022

# 2023 against 2022 deliberately. Progressive very nearly broke even in 2022 -
# 722 $m of net income on 49.6 bn of revenue - and earned five times as much the
# year after. That puts three lines past the +265% the relative tier can draw
# and one past the -35% at the other end, so the clipping notation this template
# exists to demonstrate is actually exercised by the shipped file rather than
# described in a footnote.
#
# It also breaks a rule this project's own plan wrote down: "no subtotal row is
# an outlier". That was a property of the *Institute's* figures, not a rule of
# IBCS - in their data the five outliers all happen to be elements. Here the
# subtotals are the outliers, because that is what recovering from a break-even
# year looks like. The bar clips, the triangles say it goes further, and the
# label prints the true figure; suppressing that would be misreporting to
# protect a guard.


def _xbrl_annual() -> dict[str, dict[int, float]]:
    """Annual income-statement lines, as tagged in Progressive's own 10-K.

    XBRL rather than the news releases, because a statement should come from
    the audited filing - and because these tags are the company's own, so the
    waterfall walks to the cent rather than to a tolerance.
    """
    import json

    import fetch_pgr as F

    facts = json.loads(F.fetch(
        "https://data.sec.gov/api/xbrl/companyfacts/CIK0000080661.json",
        cache_name="companyfacts.json"))["facts"]["us-gaap"]

    def series(concept: str) -> dict[int, float]:
        out: dict[int, float] = {}
        for unit in facts.get(concept, {}).get("units", {}).get("USD", []):
            start, end = unit.get("start"), unit.get("end")
            if (start and end and start[5:] == "01-01" and end[5:] == "12-31"
                    and start[:4] == end[:4]):
                out[int(end[:4])] = unit["val"] / 1e6
        return out

    return {name: series(tag) for name, tag in (
        ("revenues", "Revenues"),
        ("premiums", "PremiumsEarnedNet"),
        ("investment", "NetInvestmentIncome"),
        ("expenses", "BenefitsLossesAndExpenses"),
        ("losses", "PolicyholderBenefitsAndClaimsIncurredNet"),
        ("acquisition", "DeferredPolicyAcquisitionCostAmortizationExpense"),
        ("tax", "IncomeTaxExpenseBenefit"),
        ("net_income", "NetIncomeLoss"))}


_X = _xbrl_annual()

# (label, sign, kind, spans, how to get it). Magnitudes are stored positive and
# the sign says which way the line runs - the convention the reference uses, so
# the waterfall machinery reads them without being told anything new.
C12A_LINES = (
    ("Premiums earned", 1, "element", None,
     lambda y: _X["premiums"][y]),
    ("Investment income", 1, "element", None,
     lambda y: _X["investment"][y]),
    ("Other revenues", 1, "element", None,
     lambda y: _X["revenues"][y] - _X["premiums"][y] - _X["investment"][y]),
    ("Total revenues", 1, "subtotal", "zero", None),
    ("Losses and LAE", -1, "element", None,
     lambda y: _X["losses"][y]),
    ("Acquisition costs", -1, "element", None,
     lambda y: _X["acquisition"][y]),
    ("Other expenses", -1, "element", None,
     lambda y: _X["expenses"][y] - _X["losses"][y] - _X["acquisition"][y]),
    ("Pretax income", 1, "subtotal", "zero", None),
    ("Income taxes", -1, "element", None,
     lambda y: _X["tax"][y]),
    ("Net income", 1, "subtotal", "zero", None),
)

C12A_CATEGORIES = tuple(label for label, _s, _k, _sp, _f in C12A_LINES)
C12A_ROWS = tuple(D.Row(label=label, sign=sign, kind=kind, spans=spans)
                  for label, sign, kind, spans, _f in C12A_LINES)


def _c12a_values(year: int) -> tuple[float, ...]:
    """One year down the statement, subtotals derived from the lines above."""
    out: list[float] = []
    for (_label, sign, kind, _spans, getter) in C12A_LINES:
        if kind == "subtotal":
            running = 0.0
            for i, (_l, s, k, _sp, _g) in enumerate(C12A_LINES[:len(out)]):
                if k != "subtotal":
                    running += s * out[i]
            out.append(D.excel_round(running))
        else:
            out.append(D.excel_round(getter(year)))
    return tuple(out)


C12A_AC = _c12a_values(C12A_AC_YEAR)
C12A_PY = _c12a_values(C12A_PY_YEAR)
C12A_VAR_ABS = tuple(
    sign * (a - p) for (_l, sign, _k, _sp, _g), a, p
    in zip(C12A_LINES, C12A_AC, C12A_PY))
# None where the prior year is not positive, matching what the workbook draws.
# A percentage of a negative base gets the direction backwards - other revenues
# recovered from -866 to +1,578 and the quotient reads -282% - so the sheet
# suppresses it, and the data layer has to agree or the tie-outs would be
# checking a figure nobody sees.
C12A_VAR_REL = tuple(None if p <= 0 else v / p * 100.0
                     for v, p in zip(C12A_VAR_ABS, C12A_PY))

# Derived from the layout's own limit rather than typed, so the constant and
# the data cannot disagree about which rows overflow.
C12A_OUTLIERS = tuple(i for i, v in enumerate(C12A_VAR_REL)
                      if v is not None
                      and (v > L.C12A_REL_LIMIT[1] or v < L.C12A_REL_LIMIT[0]))

def _c12a_bounds() -> dict:
    """Axis bounds as fractions of the shared span, one tier at a time.

    The three unit tiers must share a scale, and sharing a scale means equal
    *points per unit* - so a tier's bound range has to be in the same ratio to
    its height as every other tier's. The layout declares both, so the ratio is
    read from it rather than guessed: get it wrong and the waterfalls and the
    variance beside them are drawn at different scales while every one of them
    still fits its own axis. ``test_rescale`` is what catches that, and it did.

    ``var_rel`` is deliberately absent: it is declared in percent and clips
    rather than rescales, so it is already independent of magnitude.
    """
    tiers = {spec.key: spec for spec in L.layout_for("C12A").tiers}
    wf_lo, wf_hi = -0.03, 1.0556          # the reference's own 5.6% headroom
    span = (wf_hi - wf_lo) * (tiers["var_abs"].plot_extent
                              / tiers["wf_ac"].plot_extent)
    # Placed to hold this data: 2022's near-break-even puts a larger negative
    # variance on the page than the reference ever carried.
    lo = -0.16
    return {"wf_py": (wf_lo, wf_hi), "wf_ac": (wf_lo, wf_hi),
            "var_abs": (lo, lo + span)}


_ni = C12A_CATEGORIES.index("Net income")

C12A = D.Template(
    id="C12",
    variant="A",
    kind="profit and loss waterfall",
    title=D.TitleBlock(
        entity=ENTITY,
        measure="Profit and loss statement",
        unit="$m",
        period=f"AC {C12A_AC_YEAR}, PY and ΔPY",
        message=(
            f"Net income rose from {C12A_PY[_ni]:,.0f} to {C12A_AC[_ni]:,.0f} "
            f"$m, {C12A_VAR_REL[_ni]:+,.0f}%, as an underwriting result that "
            f"had nearly vanished in {C12A_PY_YEAR} recovered; "
            f"{len(C12A_OUTLIERS)} lines move too far to be drawn to scale and "
            f"carry the overflow marks"
        ),
    ),
    categories=C12A_CATEGORIES,
    category_scenarios=("AC",) * len(C12A_CATEGORIES),
    rows=C12A_ROWS,
    tiers=(
        D.Tier(key="wf_py", label=str(C12A_PY_YEAR), kind="waterfall",
               number_format="{:,.0f}",
               series=(D.Series("PY", C12A_PY),)),
        D.Tier(key="wf_ac", label=str(C12A_AC_YEAR), kind="waterfall",
               number_format="{:,.0f}",
               series=(D.Series("AC", C12A_AC),)),
        D.Tier(key="var_abs", label="ΔPY", kind="variance_abs", reference="PY",
               number_format="{:+,.0f}",
               series=(D.Series("AC", C12A_VAR_ABS),)),
        D.Tier(key="var_rel", label="ΔPY%", kind="variance_rel", reference="PY",
               number_format="{:+,.0f}",
               series=(D.Series("AC", C12A_VAR_REL),)),
    ),
    # Fractions of the shared span, for the reason C03A carries them: the
    # layout declares its bounds in the reference's own currency units and
    # divides them by the live span, which only lands right when the span is the
    # one they were measured against.
    #
    # ``var_rel`` is deliberately absent. That tier is declared in percent and
    # **clips rather than rescales** - the -35..+265 limits are the whole point
    # of it - so it is already independent of magnitude and must not be
    # normalised, or the overflow marks would never fire.
    #
    # The waterfall keeps the reference's 5.6% headroom. The variance tier does
    # not: its bounds are fitted to these figures, because 2022's near-break-even
    # puts a bigger negative variance on the page than the reference ever had,
    # and the reference's -0.061 floor would clip a real bar.
    tier_bounds=_c12a_bounds(),
    provenance=(
        D.SourceNote("AC", "filed",
                     f"Income statement lines as tagged in Progressive's "
                     f"{C12A_AC_YEAR} 10-K (XBRL companyfacts). Subtotals are "
                     f"derived from the lines above them."),
        D.SourceNote("PY", "filed",
                     f"The same tags for {C12A_PY_YEAR}."),
    ),
    source_ref="SEC EDGAR CIK 0000080661, XBRL companyfacts",
    notes=(
        "Every subtotal is a live sum of the lines above it, so a corrected "
        "line moves the statement all the way down.",
        "Revenues less expenses less tax equals net income exactly in both "
        "years - the tags are the company's own.",
    ),
)


TEMPLATES: dict[str, D.Template] = {"C01A": C01A, "C02A": C02A,
                                    "C03A": C03A, "C10D": C10D, "C12A": C12A}
