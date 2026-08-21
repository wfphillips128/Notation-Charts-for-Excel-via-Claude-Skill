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

_FACTS: dict | None = None

# Quarter ends, as XBRL spells them. A fact is a quarter's if its period ends
# on one of these; the flows are tagged year-to-date, so a discrete quarter is
# the difference between consecutive ones.
_QUARTER_END = {"03-31": 1, "06-30": 2, "09-30": 3, "12-31": 4}


def _facts() -> dict:
    """Progressive's XBRL company facts, fetched once and parsed once.

    Three templates read from this file and each had been opening and parsing
    it for itself. ``fetch_pgr.fetch`` caches the bytes on disk so the cost was
    never the network - but three copies of the same six lines is three places
    for one of them to drift, which is exactly how one fill colour came to be
    written out fourteen times.
    """
    global _FACTS
    if _FACTS is None:
        import json

        import fetch_pgr as F

        _FACTS = json.loads(F.fetch(
            "https://data.sec.gov/api/xbrl/companyfacts/CIK0000080661.json",
            cache_name="companyfacts.json"))["facts"]["us-gaap"]
    return _FACTS


def _usd(concept: str) -> list[dict]:
    return _facts().get(concept, {}).get("units", {}).get("USD", [])


def _annual(concept: str) -> dict[int, float]:
    """One concept's full-year figures, in $m, keyed by year."""
    out: dict[int, float] = {}
    for unit in _usd(concept):
        start, end = unit.get("start"), unit.get("end")
        if (start and end and start[5:] == "01-01" and end[5:] == "12-31"
                and start[:4] == end[:4]):
            out[int(end[:4])] = unit["val"] / 1e6
    return out


def _year_end(concept: str) -> dict[int, float]:
    """One balance-sheet concept at each 31 December, in $m."""
    return {int(u["end"][:4]): u["val"] / 1e6 for u in _usd(concept)
            if u.get("end", "").endswith("12-31") and not u.get("start")}


def _quarter_end(concept: str) -> dict[tuple[int, int], float]:
    """One balance-sheet concept at each quarter end, in $m."""
    out: dict[tuple[int, int], float] = {}
    for u in _usd(concept):
        end = u.get("end", "")
        if u.get("start") or end[5:] not in _QUARTER_END:
            continue
        out[(int(end[:4]), _QUARTER_END[end[5:]])] = u["val"] / 1e6
    return out


def _quarterly(concept: str) -> dict[tuple[int, int], float]:
    """One flow concept as *discrete* quarters, in $m.

    Progressive tags its flows year to date - three months, then six, then
    nine, then twelve - so a quarter on its own is the difference between
    consecutive filings. Q1 needs no subtraction; a quarter whose predecessor
    is missing is dropped rather than guessed at, because a year-to-date figure
    silently treated as a quarter is three times too big and still plausible.
    """
    ytd: dict[tuple[int, int], float] = {}
    for u in _usd(concept):
        start, end = u.get("start"), u.get("end")
        if (not start or start[5:] != "01-01" or start[:4] != end[:4]
                or end[5:] not in _QUARTER_END):
            continue
        ytd[(int(end[:4]), _QUARTER_END[end[5:]])] = u["val"] / 1e6
    return {(y, q): v - (ytd[(y, q - 1)] if q > 1 else 0.0)
            for (y, q), v in ytd.items() if q == 1 or (y, q - 1) in ytd}


def _xbrl_annual() -> dict[str, dict[int, float]]:
    """Annual income-statement lines, as tagged in Progressive's own 10-K.

    XBRL rather than the news releases, because a statement should come from
    the audited filing - and because these tags are the company's own, so the
    waterfall walks to the cent rather than to a tolerance.
    """
    return {name: _annual(tag) for name, tag in (
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


# --------------------------------------------------------------------------- #
# C13D - fifteen years of monthly premium, as small multiples
# --------------------------------------------------------------------------- #

# **Years, not places.** The reference draws fifteen locations over twelve
# years. Progressive's 10-K names only its top ten states plus "All other", and
# the membership of that ten changes - Louisiana was in it until the mid-2010s
# and Arizona is now - so fifteen states over twelve years is a ragged grid with
# holes in exactly the years you cannot go back and fill.
#
# Fifteen *years* of twelve *months* is the same 180 figures, complete,
# rectangular, and every panel a like-for-like comparison by construction. The
# panel keys are years because that is what they are; leaving them named after
# German cities would have been a trap for the next reader.
C13D_YEARS = tuple(range(2011, 2026))
C13D_KEYS = tuple(f"y{y}" for y in C13D_YEARS)
C13D_MONTHLY = {f"y{y}": tuple(_num(_M[(y, m)]["npw"]) for m in range(1, 13))
                for y in C13D_YEARS}

# The reference series: the mean of the fifteen panels, month by month. Every
# panel is then drawn as its distance from that, which is what makes the grid
# readable - a variance in one panel means the same as a variance in the next.
C13D_AVERAGE = tuple(
    D.excel_round(sum(C13D_MONTHLY[k][m] for k in C13D_KEYS) / len(C13D_KEYS))
    for m in range(12))


def _c13d_variance(key: str) -> tuple[float, ...]:
    """One year's distance from the fifteen-year monthly mean, in percent.

    Computed as ``(v*n - total) / total`` rather than by dividing twice through
    the average - the two are the same quantity, but the second rounds twice and
    the difference shows at a half boundary. The reference makes the same point.
    """
    n = len(C13D_KEYS)
    out = []
    for m in range(12):
        total = sum(C13D_MONTHLY[k][m] for k in C13D_KEYS)
        out.append((C13D_MONTHLY[key][m] * n - total) / total * 100.0)
    return tuple(out)


C13D_VARIANCE = {k: _c13d_variance(k) for k in C13D_KEYS}

# The panel and month furthest from the mean, for the callout - found rather
# than chosen, so it cannot point at the wrong bar after a re-harvest.
_peak_key, _peak_month = max(
    ((k, m) for k in C13D_KEYS for m in range(12)),
    key=lambda km: C13D_VARIANCE[km[0]][km[1]])

_C13D_GROWTH = (sum(C13D_MONTHLY[C13D_KEYS[-1]])
                / sum(C13D_MONTHLY[C13D_KEYS[0]]))

C13D_GRID = D.PanelGrid(
    rows=4, cols=4,
    cells=tuple(D.PanelCell(key=C13D_KEYS[i] if i < len(C13D_KEYS) else None,
                            row=i // 4, col=i % 4)
                for i in range(16)),
    reference="average",
)

C13D = D.Template(
    id="C13",
    variant="D",
    kind="small multiples",
    title=D.TitleBlock(
        entity=ENTITY,
        measure="Net premiums written",
        unit="$m, relative variance in % from the monthly average",
        period=f"{C13D_YEARS[0]}..{C13D_YEARS[-1]}",
        message=(
            f"Every month of {_peak_key[1:]} sits above the fifteen-year "
            f"average, {D.MONTHS[_peak_month]} by "
            f"{C13D_VARIANCE[_peak_key][_peak_month]:+,.0f}%; the book has "
            f"grown {_C13D_GROWTH:.1f}x since {C13D_YEARS[0]}, so the grid "
            f"reads as a climb from below the line to well above it"
        ),
    ),
    categories=D.MONTHS,
    category_scenarios=("AC",) * 12,
    tiers=tuple(
        D.Tier(key=k, label=str(y), kind="variance_rel", reference="AVG",
               number_format="{:+,.0f}", printed=None,
               series=(D.Series("AC", C13D_VARIANCE[k]),))
        for k, y in zip(C13D_KEYS, C13D_YEARS)
    ) + (
        D.Tier(key="average", label=f"average {len(C13D_KEYS)} years",
               kind="measure", number_format="{:,.0f}", printed=None,
               series=(D.Series("AC", C13D_AVERAGE),)),
    ),
    panel_values=C13D_MONTHLY,
    panel_grids={"uniform": C13D_GRID},
    annotations=(D.Annotation(kind="oval", target=(_peak_key, _peak_month)),),
    provenance=(
        D.SourceNote("AC", "filed",
                     "Monthly net premiums written from the monthly 8-K "
                     "exhibits. Each year's twelve figures reconcile to the "
                     "annual total in that year's 10-K."),
    ),
    source_ref="SEC EDGAR CIK 0000080661, monthly 8-K exhibits",
    notes=(
        "Fifteen panels on one scale - a variance read in one panel means the "
        "same in the next, which is the whole argument of a small multiple.",
        "Panels are years and categories are months, because the geographic "
        "cut Progressive publishes cannot fill a rectangle.",
    ),
)


# --------------------------------------------------------------------------- #
# C04A - underwriting result by month against the 96 combined ratio target
# --------------------------------------------------------------------------- #

# The first template here with a **plan**, and the only kind of plan a public
# company gives you. Progressive states a 96 combined ratio target - a 4%
# underwriting margin - in every 10-K, and it drives the Gainsharing programme.
# The target is filed. Applying it to a month to get a plan figure is this
# workbook's construction, so the PL scenario is marked ``assumed`` and the
# sheet shades and footnotes it.
#
# 2023 rather than a better year on purpose: five months fell short of the
# target and seven beat it, so the variance tier carries both colours. In 2025
# only one month missed, and a variance chart where everything points one way
# demonstrates nothing.
C04A_YEAR = 2023


def _c04a_rows() -> list[tuple[str, float, float]]:
    """(month, underwriting result, plan) for the year, sorted by the gap.

    Sorted by ΔPL descending, because that ordering *is* variant A - the
    template is a ranking, and re-sorting it is not a presentation choice.
    """
    out = []
    for m in range(1, 13):
        row = _M[(C04A_YEAR, m)]
        npe = float(row["npe"])
        result = npe * (100.0 - float(row["combined_ratio"])) / 100.0
        plan = npe * TARGET_MARGIN / 100.0
        out.append((D.MONTHS[m - 1], result, plan))
    return sorted(out, key=lambda r: -(r[1] - r[2]))


_c04a = _c04a_rows()
C04A_CATEGORIES = tuple(r[0] for r in _c04a)
C04A_MEASURE = tuple(D.excel_round(r[1]) for r in _c04a)
C04A_PLAN = tuple(D.excel_round(r[2]) for r in _c04a)
C04A_VAR_ABS = tuple(a - p for a, p in zip(C04A_MEASURE, C04A_PLAN))
C04A_VAR_REL = tuple(v / p * 100.0 for v, p in zip(C04A_VAR_ABS, C04A_PLAN))

_beat = sum(1 for v in C04A_VAR_ABS if v > 0)

C04A = D.Template(
    id="C04",
    variant="A",
    kind="multi-tier bar chart",
    orientation="horizontal",
    title=D.TitleBlock(
        entity=ENTITY,
        measure="Underwriting result",
        unit="$m",
        period=str(C04A_YEAR),
        message=(
            f"{_beat} of 12 months beat the 96 combined ratio target in "
            f"{C04A_YEAR}; {C04A_CATEGORIES[0]} by "
            f"{C04A_VAR_ABS[0]:+,.0f} $m and {C04A_CATEGORIES[-1]} missed it "
            f"by {abs(C04A_VAR_ABS[-1]):,.0f} $m, a spread of "
            f"{C04A_VAR_ABS[0] - C04A_VAR_ABS[-1]:,.0f} $m across the year"
        ),
    ),
    categories=C04A_CATEGORIES,
    category_scenarios=("AC",) * 12,
    tiers=(
        D.Tier(key="measure", label="", kind="measure", number_format="{:,.0f}",
               series=(D.Series("PL", C04A_PLAN, label="PL"),
                       D.Series("AC", C04A_MEASURE, label="AC"))),
        D.Tier(key="var_abs", label="ΔPL", kind="variance_abs", reference="PL",
               number_format="{:+,.0f}", printed=None,
               series=(D.Series("AC", C04A_VAR_ABS),)),
        D.Tier(key="var_rel", label="ΔPL%", kind="variance_rel", reference="PL",
               number_format="{:+,.0f}", printed=None,
               series=(D.Series("AC", C04A_VAR_REL),)),
    ),
    # Fractions of each span. The measure tier carries more headroom than it
    # needs, and that is the shared scale rather than a choice: the unit tiers
    # must be drawn at the same points per unit, so the measure range is fixed
    # by the variance range and the two tier heights. Progressive's plan is 4%
    # of premium and its result was two to four times that, which makes the
    # variance nearly as large as the measure - where the reference's variance
    # is a fraction of its net sales.
    tier_bounds={
        "measure": (-0.35, 1.826),
        "var_abs": (-0.55, 0.81),
        "var_rel": (-0.95, 1.15),
    },
    provenance=(
        D.SourceNote("AC", "filed",
                     f"Monthly net premiums earned and combined ratio from the "
                     f"{C04A_YEAR} monthly 8-K exhibits; the underwriting "
                     f"result is premium times (100 - combined ratio)."),
        D.SourceNote("PL", "assumed",
                     f"Progressive publishes no monthly plan. The target it "
                     f"does publish is a 96 combined ratio - a "
                     f"{TARGET_MARGIN:.0f}% underwriting margin - applied here "
                     f"to each month's earned premium."),
    ),
    source_ref="SEC EDGAR CIK 0000080661, monthly 8-K exhibits",
    notes=(
        "Sorted by the gap to plan, largest first. That ordering is what "
        "makes this variant A.",
        "The plan is constructed, not reported: its cells are shaded "
        "differently and the basis is stated under the data.",
    ),
)


# --------------------------------------------------------------------------- #
# C09C - earned premium against margin, with iso-result curves
# --------------------------------------------------------------------------- #

# The template's own arithmetic, arriving intact from a different industry:
# gross profit is net sales times margin, so a constant gross profit is a
# hyperbola and the chart draws curves rather than a line. An insurer's
# underwriting result is earned premium times underwriting margin - the same
# product of the same two axes. Nothing is bent to fit.
C09C_SEGMENTS = ("agency", "direct", "commercial")
C09C_LABELS = {"agency": "Agency auto", "direct": "Direct auto",
               "commercial": "Commercial"}
C09C_YEARS = (2024, 2025)

# Property is left out, and not to tidy the picture. It is small enough that a
# single catastrophe month takes its combined ratio past 280 - a margin of
# -183% - and the iso-curves live in the positive quadrant. A point at -183
# would not be drawn at all, which is worse than not claiming to draw it.
_C09C_EXCLUDED = "property"


def _c09c_points() -> tuple[D.Point, ...]:
    rows = [r for r in _rows("monthly_segments.csv")
            if r["role"] in C09C_SEGMENTS and r["npe"] and r["combined_ratio"]
            and int(r["year"]) in C09C_YEARS]
    by_month: dict[tuple[int, int], dict] = {}
    for r in rows:
        by_month.setdefault((int(r["year"]), int(r["month"])), {})[r["role"]] = r

    out: list[D.Point] = []
    for key in sorted(k for k, v in by_month.items()
                      if len(v) == len(C09C_SEGMENTS)):
        for role in C09C_SEGMENTS:
            r = by_month[key][role]
            out.append(D.Point(
                # Unnamed, like the reference's own points: sixty-nine labels
                # on one scattergram is not a chart, and the month a point
                # belongs to is in the source rows beside it.
                entity="",
                scenario="AC",
                x=100.0 - float(r["combined_ratio"]),
                y=float(r["npe"]),
                group=C09C_LABELS[role],
            ))
    return tuple(out)


C09C_POINTS = _c09c_points()
C09C_LEVELS = (100.0, 200.0, 300.0)
C09C_SEGMENT = C09C_LEVELS[-1]
_c09c_counts = {
    label: sum(1 for p in C09C_POINTS
               if p.group == label and p.x * p.y / 100.0 >= C09C_SEGMENT)
    for label in C09C_LABELS.values()}
C09C_MESSAGE_GROUP = max(_c09c_counts, key=_c09c_counts.get)
C09C_COUNT = _c09c_counts[C09C_MESSAGE_GROUP]
_c09c_months = len(C09C_POINTS) // len(C09C_SEGMENTS)

C09C = D.Template(
    id="C09",
    variant="C",
    kind="scattergram",
    title=D.TitleBlock(
        entity=ENTITY,
        measure="Net premiums earned",
        unit="$m, margin in %",
        period=f"{C09C_YEARS[0]}..{C09C_YEARS[-1]}",
        message=(
            f"{C09C_COUNT} of {C09C_MESSAGE_GROUP}'s {_c09c_months} months "
            f"earned an underwriting result of {C09C_SEGMENT:,.0f} $m or more; "
            f"the curves are constant result, so a point's distance beyond one "
            f"is what its premium and its margin achieved together"
        ),
        subject=(("Net premiums earned", "bold"), (" in $m, ", "normal"),
                 ("margin", "bold"), (" in %", "normal")),
    ),
    categories=tuple(p.entity for p in C09C_POINTS),
    category_scenarios=("AC",) * len(C09C_POINTS),
    tiers=(),
    points=C09C_POINTS,
    axes=(D.Axis("Underwriting margin", 0.0, 25.0, 5.0, "{:.0f}",
                 unit="% of earned premium"),
          D.Axis("Net premiums earned", 0.0, 3500.0, 500.0, "{:,.0f}",
                 unit="$m")),
    groups=tuple(D.Group(label, i)
                 for i, label in enumerate(C09C_LABELS.values())),
    iso_curves=D.IsoCurves(levels=C09C_LEVELS, segment=C09C_SEGMENT,
                           group=C09C_MESSAGE_GROUP, count=C09C_COUNT),
    provenance=(
        D.SourceNote("AC", "filed",
                     "Monthly earned premium and combined ratio by segment "
                     "from the monthly 8-K exhibits; margin is 100 less the "
                     "combined ratio, and the result is their product."),
    ),
    source_ref="SEC EDGAR CIK 0000080661, monthly 8-K exhibits",
    notes=(
        "Colour carries the segment, which is a category and not a scenario - "
        "the accent ramp, never the scenario greys.",
        "Property is excluded: one catastrophe month takes its margin to -183%, "
        "and the iso-curves live in the positive quadrant.",
    ),
)


# --------------------------------------------------------------------------- #
# C11A - the return tree: margin x turnover
# --------------------------------------------------------------------------- #

# The reference's tree is return on investment decomposed into return on sales
# and capital turnover. An insurer's is the same decomposition of the same
# quantities: underwriting result over earned premium is the margin, earned
# premium over shareholders' equity is the turnover, and their product is the
# return on that equity. The links hold to the digit.
C11A_YEARS = tuple(range(2019, 2026))


def _c11a_series() -> dict[str, tuple[float, ...]]:
    equity = _year_end("StockholdersEquity")

    seg = annual_segments()
    npe, result, capital = [], [], []
    for y in C11A_YEARS:
        row = seg[y]["companywide"]
        earned = float(row["npe"])
        npe.append(earned)
        result.append(earned * (100.0 - float(row["combined_ratio"])) / 100.0)
        capital.append(equity[y])
    # Round the three base measures *first*, then derive the ratios from the
    # rounded figures. The workbook computes its quotients as formulas over the
    # cells it displays, so deriving here from unrounded values would put a
    # tree on the page whose own arithmetic did not quite close - by a
    # thousandth, which is exactly the kind of discrepancy nobody finds by
    # looking.
    result = [D.excel_round(v) for v in result]
    npe = [D.excel_round(v) for v in npe]
    capital = [D.excel_round(v) for v in capital]
    ros = [r / n * 100.0 for r, n in zip(result, npe)]
    turn = [n / c for n, c in zip(npe, capital)]
    return {
        "return": tuple(result),
        "net_sales": tuple(npe),
        "capital": tuple(capital),
        "ros": tuple(ros), "turnover": tuple(turn),
        "roi": tuple(a * b for a, b in zip(ros, turn)),
    }


C11A_VALUES = _c11a_series()


def _c11a_scale_px(tallest: float = 240.0) -> dict:
    """Pixels per unit for each scale group, from this data's own ranges.

    One scale per unit, never one per box - six boxes on three rulers is what
    makes the tree read as a single picture. The rulers are sized so the tallest
    box in each group lands near the height the reference gives it; the others
    then come out shorter or taller in proportion, which is the consequence of
    sharing a scale rather than a fault in it.

    Declared in pixels because the SVG engine draws in pixels and the workbook
    converts. Not inherited from the layout: those constants were measured
    against figures in the hundreds, and premiums in the tens of thousands would
    draw a box some hundreds of feet tall.
    """
    groups = {"percent": ("roi", "ros"), "kEUR": ("return", "net_sales",
                                                  "capital"),
              "turnover": ("turnover",)}
    out = {}
    for group, keys in groups.items():
        widest = max(max(C11A_VALUES[k]) - min(C11A_VALUES[k]) for k in keys)
        out[group] = tallest / widest / 0.75      # points -> pixels
    return out


C11A_TREE = D.TreeSpec(
    scale_px=_c11a_scale_px(),
    nodes=(
        D.TreeNode(key="roi", column=0, scale_group="percent", unit="%"),
        D.TreeNode(key="ros", column=1, scale_group="percent", unit="%"),
        D.TreeNode(key="turnover", column=1, scale_group="turnover"),
        D.TreeNode(key="return", column=2, scale_group="kEUR", unit="$m"),
        D.TreeNode(key="net_sales", column=2, scale_group="kEUR", unit="$m"),
        D.TreeNode(key="capital", column=2, scale_group="kEUR", unit="$m"),
    ),
    links=(
        D.TreeLink(operator=":", left="return", right="net_sales", result="ros"),
        D.TreeLink(operator=":", left="net_sales", right="capital",
                   result="turnover"),
        D.TreeLink(operator="x", left="ros", right="turnover", result="roi"),
    ),
    printed_categories=(0, len(C11A_YEARS) - 1),
)

_c11a_labels = {"roi": "Return on equity", "ros": "Underwriting margin",
                "turnover": "Premium to equity", "return": "Underwriting result",
                "net_sales": "Net premiums earned", "capital": "Shareholders' equity"}
_lo, _hi = C11A_VALUES["roi"][0], C11A_VALUES["roi"][-1]

C11A = D.Template(
    id="C11",
    variant="A",
    kind="tree chart",
    title=D.TitleBlock(
        entity=ENTITY,
        measure="Return on equity",
        unit="%",
        period=f"{C11A_YEARS[0]}..{C11A_YEARS[-1]}",
        message=(
            f"Return on equity moved from {_lo:.1f}% to {_hi:.1f}% between "
            f"{C11A_YEARS[0]} and {C11A_YEARS[-1]}, and the tree says which "
            f"half did it: premium to equity barely moved "
            f"({C11A_VALUES['turnover'][0]:.2f} to "
            f"{C11A_VALUES['turnover'][-1]:.2f}) while the underwriting "
            f"margin went from {C11A_VALUES['ros'][0]:.1f}% to "
            f"{C11A_VALUES['ros'][-1]:.1f}%"
        ),
    ),
    categories=tuple(str(y) for y in C11A_YEARS),
    category_scenarios=("AC",) * len(C11A_YEARS),
    tiers=tuple(
        D.Tier(key=k, label=_c11a_labels[k], kind="measure",
               number_format="{:,.1f}" if k in ("roi", "ros") else
                             ("{:,.2f}" if k == "turnover" else "{:,.0f}"),
               printed=None,
               series=(D.Series("AC", C11A_VALUES[k]),))
        for k in ("roi", "ros", "turnover", "return", "net_sales", "capital")),
    tree=C11A_TREE,
    provenance=(
        D.SourceNote("AC", "filed",
                     "Earned premium and combined ratio from the year-to-date "
                     "block of each December release; shareholders' equity as "
                     "tagged in the 10-K. Margin, turnover and return on "
                     "equity are their quotients and product."),
    ),
    source_ref="SEC EDGAR CIK 0000080661, monthly 8-K exhibits and XBRL",
    notes=(
        "One scale per unit, not one per box: six boxes on three rulers is "
        "what makes the tree one picture.",
        "The rulers are sized from these figures. The layout's own are "
        "measured against a reference in the hundreds.",
    ),
)


# --------------------------------------------------------------------------- #
# C08H - the stock roll-forward: loss and loss adjustment expense reserves      #
# --------------------------------------------------------------------------- #
#
# The template's law is ``level = opening + sum(flow - outflow)``, which is
# exactly how an insurer's claim reserve moves: what it books as incurred each
# quarter goes in, what it pays out comes off, and the balance is whatever is
# still owed on claims that have not been settled. Nothing was bent to fit -
# the roll-forward is a table Progressive prints in every 10-K and 10-Q.
#
# Quarterly rather than annual because the flows are tagged year to date at
# each quarter end, so differencing them yields discrete quarters - and twenty
# of them, which is the number of categories the template carries.

C08H_ACTUALS = ([(2024, q) for q in (1, 2, 3, 4)]
                + [(2025, q) for q in (1, 2, 3, 4)]
                + [(2026, 1), (2026, 2)])
C08H_FORECAST = [(2026, 3), (2026, 4)]
C08H_PLAN = [(y, q) for y in (2027, 2028) for q in (1, 2, 3, 4)]
C08H_QUARTERS = C08H_ACTUALS + C08H_FORECAST + C08H_PLAN
C08H_SCENARIOS = (("AC",) * len(C08H_ACTUALS) + ("FC",) * len(C08H_FORECAST)
                  + ("PL",) * len(C08H_PLAN))

# The roll-forward is stated *net* of reinsurance, so every figure in it has to
# be the net one. The gross balance sits in a different tag and does not close
# against these flows - it is larger by the reinsurance recoverable, which is
# real money but not Progressive's to pay.
C08H_TAGS = {
    "balance": "LiabilityForUnpaidClaimsAndClaimsAdjustmentExpenseNet",
    "incurred":
        "LiabilityForUnpaidClaimsAndClaimsAdjustmentExpenseIncurredClaims1",
    "paid_current":
        "LiabilityForUnpaidClaimsAndClaims"
        "AdjustmentExpenseClaimsPaidCurrentYear1",
    "paid_prior":
        "LiabilityForUnpaidClaimsAndClaims"
        "AdjustmentExpenseClaimsPaidPriorYears1",
}


def _c08h_series() -> dict:
    """The reserve roll-forward: ten quarters filed, ten constructed.

    The filed quarters are read straight out of the filings, and the derived
    level is checked against the balance Progressive reports at each quarter
    end rather than merely being self-consistent: ``opening + sum of changes``
    has to arrive at the number in the filing, quarter by quarter.

    The constructed quarters follow the workbook's one plan basis - the 96
    combined ratio Progressive states in every 10-K - resolved into a reserve
    movement by three rules, each of which is printed on the sheet:

    * **Earned premium** continues the trailing four quarters' year-on-year
      growth, compounded quarterly. Held flat instead, every constructed
      quarter comes out identical, and ten identical bars read as a file
      nobody filled in.
    * **Losses incurred** is that premium at the loss ratio the target implies:
      96 less the trailing four quarters' expense ratio. The target is filed
      and the expense ratio is filed; only applying them to a future quarter is
      ours.
    * **Claims paid** is incurred at the trailing four quarters' ratio of
      payments to losses incurred, which preserves the payment lag that is what
      makes a reserve a reserve.
    """
    balance = _quarter_end(C08H_TAGS["balance"])
    incurred = _quarterly(C08H_TAGS["incurred"])
    paid_cur = _quarterly(C08H_TAGS["paid_current"])
    paid_pri = _quarterly(C08H_TAGS["paid_prior"])
    paid = {q: paid_cur[q] + paid_pri[q] for q in paid_cur if q in paid_pri}

    earned = _quarterly("PremiumsEarnedNet")
    losses = _quarterly("PolicyholderBenefitsAndClaimsIncurredNet")
    outgo = _quarterly("BenefitsLossesAndExpenses")

    missing = [q for q in C08H_ACTUALS
               if q not in incurred or q not in paid or q not in balance]
    if missing:
        raise ValueError(
            f"C08H: no filed reserve roll-forward for {missing} - the window "
            f"in C08H_ACTUALS reaches past what Progressive has filed")

    first = C08H_ACTUALS[0]
    before = (first[0] - 1, 4) if first[1] == 1 else (first[0], first[1] - 1)
    opening = balance[before]

    flow = [incurred[q] for q in C08H_ACTUALS]
    outflow = [paid[q] for q in C08H_ACTUALS]
    filed_levels = [balance[q] for q in C08H_ACTUALS]

    # The trailing four quarters set every rate the constructed side uses, so
    # they are read once and named rather than recomputed at each use.
    tail = C08H_ACTUALS[-4:]
    prior = C08H_ACTUALS[-8:-4]
    tail_earned = sum(earned[q] for q in tail)
    tail_losses = sum(losses[q] for q in tail)
    tail_outgo = sum(outgo[q] for q in tail)
    expense_ratio = (tail_outgo - tail_losses) / tail_earned * 100.0
    plan_loss_ratio = TARGET_COMBINED_RATIO - expense_ratio
    pay_rate = sum(outflow[-4:]) / sum(flow[-4:])
    growth = (tail_earned / sum(earned[q] for q in prior)) ** 0.25

    premium = earned[C08H_ACTUALS[-1]]
    for _quarter in C08H_FORECAST + C08H_PLAN:
        premium *= growth
        # Rounded here, before the payment is taken off it, so that the two
        # figures the sheet prints reproduce the rate the sheet states. Derive
        # the payment from the unrounded booking instead and a reader dividing
        # the printed pair gets a slightly different percentage from the one in
        # the footnote - the same defect the return tree had.
        booked = D.excel_round(premium * plan_loss_ratio / 100.0)
        flow.append(booked)
        outflow.append(booked * pay_rate)

    # Round the two flows first, then derive the change and the level from the
    # rounded figures - the lesson the return tree taught. The workbook
    # recomputes its own levels as formulas over the cells it prints, so a
    # level derived here from unrounded flows would disagree with the one the
    # sheet draws, by a hair and invisibly.
    flow = [D.excel_round(v) for v in flow]
    outflow = [D.excel_round(v) for v in outflow]
    change = [i - o for i, o in zip(flow, outflow)]
    level, running = [], D.excel_round(opening)
    for delta in change:
        running += delta
        level.append(running)

    return {
        "opening": D.excel_round(opening),
        "flow": tuple(flow), "outflow": tuple(outflow),
        "change": tuple(change), "level": tuple(level),
        "filed_levels": tuple(filed_levels),
        "expense_ratio": expense_ratio,
        "plan_loss_ratio": plan_loss_ratio,
        "actual_loss_ratio": tail_losses / tail_earned * 100.0,
        "pay_rate": pay_rate, "growth": growth,
    }


C08H_VALUES = _c08h_series()


def _c08h_split(values, keep):
    return [v if s == keep else None for v, s in zip(values, C08H_SCENARIOS)]


def _c08h_tier(key: str, label: str) -> D.Tier:
    values = C08H_VALUES[key]
    return D.Tier(
        key=key, label=label, kind="measure", height_weight=1.0,
        number_format="{:+,.0f}" if key == "change" else "{:,.0f}",
        printed=None,
        series=tuple(D.Series(scenario, _c08h_split(values, scenario))
                     for scenario in ("AC", "FC", "PL")),
    )


_c08h_last = len(C08H_ACTUALS) - 1
_c08h_built = C08H_VALUES["level"][_c08h_last] - C08H_VALUES["opening"]
_c08h_up = sum(1 for c in C08H_VALUES["change"][:len(C08H_ACTUALS)] if c > 0)

C08H = D.Template(
    id="C08",
    variant="H",
    kind="area chart",
    title=D.TitleBlock(
        entity=ENTITY,
        measure="Loss and loss adjustment expense reserves, net",
        unit="$m",
        period=f"{C08H_QUARTERS[0][0]}..{C08H_QUARTERS[-1][0]}",
        message=(
            f"Reserves built {_c08h_built:+,.0f} $m over "
            f"{len(C08H_ACTUALS)} quarters - losses incurred ran ahead of "
            f"claims paid in {_c08h_up} of them. Progressive is running a "
            f"{C08H_VALUES['actual_loss_ratio']:.1f} loss ratio where its "
            f"{TARGET_COMBINED_RATIO:.0f} combined ratio target implies "
            f"{C08H_VALUES['plan_loss_ratio']:.1f}, so holding the target "
            f"builds reserves faster rather than slower: "
            f"{C08H_VALUES['level'][-1]:,.0f} $m by end "
            f"{C08H_QUARTERS[-1][0]}"
        ),
    ),
    categories=tuple(f"Q{q} {y}" for y, q in C08H_QUARTERS),
    category_scenarios=C08H_SCENARIOS,
    tiers=(
        _c08h_tier("change", "Reserve change"),
        _c08h_tier("level", "Reserves, net"),
        _c08h_tier("flow", "Losses incurred"),
        _c08h_tier("outflow", "Claims paid"),
    ),
    opening=C08H_VALUES["opening"],
    # The reference circles its one forecast quarter. Here the point worth
    # circling is the handover: the last level anyone has filed, after which
    # every figure on the sheet is this workbook's construction.
    annotations=(D.Annotation(kind="oval", target=("level", _c08h_last)),),
    # Fractions of the span, not data units - the sixth instance of the same
    # defect, and the one that would have been hardest to see. ``LineLayout``
    # declares its axis maximum as 24, which is right above an inventory that
    # tops out at 22 tons and is 24/54,000ths of the way up a loss reserve, so
    # every series would clamp to the frame and draw flat.
    #
    # These are the recreation's own bounds over the recreation's own span, so
    # the picture keeps the proportions the reference render has: about a tenth
    # of headroom above, and a third of the range below zero for the payment
    # columns to hang into.
    tier_bounds={"level": (-8.0 / 22.0, 24.0 / 22.0)},
    provenance=(
        D.SourceNote("AC", "filed",
                     "Reserve roll-forward, 10-Q and 10-K. Flows are tagged "
                     "year to date and differenced to quarters; the derived "
                     "level is checked against the balance Progressive "
                     "reports at each quarter end."),
        D.SourceNote("FC", "assumed",
                     f"Progressive publishes no forecast. Earned premium "
                     f"continues the trailing four quarters' growth "
                     f"({C08H_VALUES['growth'] ** 4 * 100 - 100:+.1f}% a "
                     f"year); losses incurred are that premium at "
                     f"{C08H_VALUES['plan_loss_ratio']:.1f}%, the loss ratio "
                     f"a {TARGET_COMBINED_RATIO:.0f} combined ratio implies "
                     f"after the trailing "
                     f"{C08H_VALUES['expense_ratio']:.1f}% expense ratio; "
                     f"claims paid are {C08H_VALUES['pay_rate'] * 100:.1f}% "
                     f"of losses incurred, the trailing rate."),
        D.SourceNote("PL", "assumed",
                     "The same three rules carried through the two plan "
                     "years. The 96 combined ratio target is Progressive's "
                     "own and is filed; applying it to a future quarter is "
                     "this workbook's construction."),
    ),
    source_ref="SEC EDGAR CIK 0000080661, 10-Q and 10-K reserve roll-forward",
    notes=(
        "An opening balance and two reported flows determine the whole chart; "
        "the printed levels are the check, and the ten filed ones agree with "
        "the balance sheet to the million.",
        "The reserve is net of reinsurance throughout, which is the basis the "
        "roll-forward itself is stated on.",
        "Losses incurred here is the same figure the income statement carries "
        "as losses and loss adjustment expense - two different tables in the "
        "same filing, and they agree.",
    ),
)


# --------------------------------------------------------------------------- #
# C07C - cumulative underwriting result against the target, with a moving       #
#        annual total                                                           #
# --------------------------------------------------------------------------- #
#
# The same year C03A draws, asking a different question. C03A asks whether the
# premium is growing; this asks whether the year is earning what it set out to,
# and carries a twelve-month rolling total so the seasonality does not read as
# a trend.
#
# The reference notes that its moving annual total "is the only series that
# cannot be derived from the page - it reaches into 2024 - so it is
# transcribed". Here it *can* be derived, because the harvest holds every month
# back to 2008: the total is computed from the eleven months before each one.

C07C_YEAR = 2026
C07C_LAST_ACTUAL = 7


def _c07c_result(key: tuple[int, int]) -> float | None:
    """One month's underwriting result: earned premium at its own margin.

    An insurer's combined ratio is what it spends per dollar earned, so
    100 less the ratio is the margin and the product is the result. Both
    figures come off the same line of the same release.
    """
    row = monthly().get(key)
    if not row:
        return None
    earned, ratio = _num(row.get("npe")), _num(row.get("combined_ratio"))
    if earned is None or ratio is None:
        return None
    return earned * (100.0 - ratio) / 100.0


def _c07c_series() -> dict:
    """Twelve months of underwriting result: seven filed, five constructed.

    The constructed months use the same two rules C03A uses, for the same
    reason - a second basis for the same year would be a second thing to
    explain:

    * **Earned premium** is the prior year's month at the pace the year to date
      is running against the same months last year.
    * **The margin** is the year-to-date margin, so the rest of the year is
      assumed to run as the year so far has rather than to a fresh assumption.

    The plan is the workbook's one basis - a 96 combined ratio, so a 4.0%
    margin - applied to the premium each month actually earned, which is the
    comparison the target is meant to support: what the month *would* have
    returned at target on the premium it did earn.
    """
    year, last = C07C_YEAR, C07C_LAST_ACTUAL
    actual_months = range(1, last + 1)

    earned_ytd = sum(_num(monthly()[(year, m)]["npe"]) for m in actual_months)
    prior_ytd = sum(_num(monthly()[(year - 1, m)]["npe"]) for m in actual_months)
    pace = earned_ytd / prior_ytd
    result_ytd = sum(_c07c_result((year, m)) for m in actual_months)
    ytd_margin = result_ytd / earned_ytd * 100.0

    earned: dict[int, float] = {}
    result: dict[int, float] = {}
    for m in range(1, 13):
        if m <= last:
            earned[m] = _num(monthly()[(year, m)]["npe"])
            result[m] = _c07c_result((year, m))
        else:
            earned[m] = _num(monthly()[(year - 1, m)]["npe"]) * pace
            result[m] = earned[m] * ytd_margin / 100.0

    # Round before anything is summed. The workbook builds its cumulatives as
    # running sums of the cells it prints, so a total accumulated here from
    # unrounded months would drift from the one the sheet draws.
    result = {m: D.excel_round(v) for m, v in result.items()}
    plan = {m: D.excel_round(earned[m] * TARGET_MARGIN / 100.0)
            for m in range(1, 13)}

    def running(values: dict[int, float]) -> list[float]:
        out, total = [], 0.0
        for m in range(1, 13):
            total += values[m]
            out.append(total)
        return out

    # The moving annual total: this month and the eleven before it, reaching
    # back into the prior year wherever it has to.
    mat: list[float] = []
    for m in range(1, 13):
        total = 0.0
        for back in range(12):
            mm, yy = m - back, year
            while mm < 1:
                mm, yy = mm + 12, yy - 1
            total += result[mm] if yy == year else _c07c_result((yy, mm))
        mat.append(D.excel_round(total))

    return {
        "monthly": tuple(result[m] for m in range(1, 13)),
        "plan": tuple(plan[m] for m in range(1, 13)),
        "cum": tuple(running(result)), "cum_pl": tuple(running(plan)),
        "mat": tuple(mat),
        "pace": pace, "ytd_margin": ytd_margin,
        "earned": tuple(earned[m] for m in range(1, 13)),
    }


C07C_VALUES = _c07c_series()
C07C_SCENARIOS = (("AC",) * C07C_LAST_ACTUAL
                  + ("FC",) * (12 - C07C_LAST_ACTUAL))


def _c07c_split(values, keep):
    return [v if s == keep else None for v, s in zip(values, C07C_SCENARIOS)]


_c07c_at_date = (C07C_VALUES["cum"][C07C_LAST_ACTUAL - 1]
                 - C07C_VALUES["cum_pl"][C07C_LAST_ACTUAL - 1])
_c07c_full_year = C07C_VALUES["cum"][-1] - C07C_VALUES["cum_pl"][-1]

C07C = D.Template(
    id="C07",
    variant="C",
    kind="line chart",
    title=D.TitleBlock(
        entity=ENTITY,
        measure="Underwriting result",
        unit="$m",
        period=str(C07C_YEAR),
        message=(
            f"Through {D.MONTHS[C07C_LAST_ACTUAL - 1]} the underwriting result "
            f"is {_c07c_at_date:+,.0f} $m against the "
            f"{TARGET_COMBINED_RATIO:.0f} combined ratio target - a "
            f"{C07C_VALUES['ytd_margin']:.1f}% margin where the target implies "
            f"{TARGET_MARGIN:.1f}%. Holding that margin closes the year "
            f"{_c07c_full_year:+,.0f} $m ahead of target, and the moving annual "
            f"total has stayed above "
            f"{min(C07C_VALUES['mat']) // 1000 * 1000:,.0f} $m all year"
        ),
    ),
    categories=D.MONTHS,
    category_scenarios=C07C_SCENARIOS,
    tiers=(
        D.Tier(key="mat", label="MAT", kind="measure", printed=None,
               number_format="{:,.0f}",
               series=(D.Series("AC", _c07c_split(C07C_VALUES["mat"], "AC"),
                                label="MAT"),
                       D.Series("FC", _c07c_split(C07C_VALUES["mat"], "FC"),
                                label="MAT"))),
        D.Tier(key="cum_pl", label="PL", kind="measure", printed=None,
               number_format="{:,.0f}",
               series=(D.Series("PL", C07C_VALUES["cum_pl"], label="PL"),)),
        D.Tier(key="cum_ac", label="AC", kind="measure", printed=None,
               number_format="{:,.0f}",
               series=(D.Series("AC", _c07c_split(C07C_VALUES["cum"], "AC"),
                                label="AC"),)),
        D.Tier(key="cum_fc", label="FC", kind="measure", printed=None,
               number_format="{:,.0f}",
               series=(D.Series("FC", _c07c_split(C07C_VALUES["cum"], "FC"),
                                label="FC"),)),
        D.Tier(key="monthly", label="", kind="measure", printed=None,
               number_format="{:,.0f}",
               series=(D.Series("PL", C07C_VALUES["plan"], label="PL"),
                       D.Series("AC", _c07c_split(C07C_VALUES["monthly"], "AC"),
                                label="AC"),
                       D.Series("FC", _c07c_split(C07C_VALUES["monthly"], "FC"),
                                label="FC"))),
    ),
    # Same shape as the reference's: one at the reporting date and one at the
    # year end, each stating the gap to plan. Derived from the series rather
    # than typed, so a retyped month moves the callout and the sentence is what
    # has to be brought back into line.
    #
    # Carried because the SVG engine draws them. The Excel engine honours
    # annotations on the driver tree only, so the workbook does not yet show
    # these - a gap it has with the recreation too, not one this data creates.
    annotations=(
        D.Annotation(kind="bracket", target=("cum", C07C_LAST_ACTUAL - 1),
                     text=f"{_c07c_at_date:+,.0f}"),
        D.Annotation(kind="bracket", target=("cum", 11),
                     text=f"{_c07c_full_year:+,.0f}"),
    ),
    # Fractions of the span. ``LineLayout.maximum`` defaults to 2,200, which
    # suits a reference topping out at 2,118 kUSD and not a result in the tens
    # of thousands. These are the recreation's own bound over its own span, so
    # the headroom above the tallest series is the proportion the reference
    # render has.
    tier_bounds={"line": (0.0, 2200.0 / 2118.0)},
    provenance=(
        D.SourceNote("AC", "derived",
                     "Earned premium and the combined ratio are both from the "
                     "monthly 8-K exhibit; the result is premium times "
                     "(100 - ratio), which is the definition of the ratio "
                     "rather than an assumption."),
        D.SourceNote("PL", "assumed",
                     f"The {TARGET_COMBINED_RATIO:.0f} combined ratio target "
                     f"is Progressive's own and is stated in every 10-K, so a "
                     f"{TARGET_MARGIN:.1f}% margin is filed. Applying it month "
                     f"by month to produce a plan is this workbook's "
                     f"construction, and the target is a floor the company "
                     f"aims to beat rather than a forecast of what it expects."),
        D.SourceNote("FC", "assumed",
                     f"Remaining months are the prior year's months at the "
                     f"year-to-date pace "
                     f"({C07C_VALUES['pace'] * 100 - 100:+.1f}%), earning the "
                     f"year-to-date margin "
                     f"({C07C_VALUES['ytd_margin']:.1f}%)."),
    ),
    source_ref="SEC EDGAR CIK 0000080661, monthly 8-K exhibits",
    notes=(
        "Four series on one scale: a moving annual total, and the plan, actual "
        "and forecast cumulatives the monthly tier sums to.",
        "The moving annual total is derived here rather than transcribed - the "
        "harvest reaches back to 2008, so the eleven months before January are "
        "on hand. It meets the cumulative exactly at December, which is the "
        "one month where the two mean the same thing.",
        "Plan is the target margin on the premium each month actually earned, "
        "so the gap is margin against target and never a premium miss.",
    ),
)


# --------------------------------------------------------------------------- #
# C06F - where the year's growth came from, state by state                      #
# --------------------------------------------------------------------------- #
#
# A bar waterfall bridging one year's premium to the next, one state at a time.
# Variant F is the *ordering* - largest contribution to smallest - so the sort
# is asserted rather than assumed, exactly as the reference's own note says.
#
# The catch-all is pinned last whatever its size, which here matters more than
# it does in the reference: "All other" is the largest single contributor, and
# sorting it into first place would say the states outside the top ten are one
# place rather than forty.


def _walk_bounds(levels, span: float, pad: float = 0.15) -> tuple[float, float]:
    """The window a bridge walks through, as fractions of the sheet's span.

    A seventh instance of the recurring defect, and the one that does not yield
    to either of the usual fixes. The layout declares this window as a fixed
    band in data units - 1,600 to 2,200 - which is wrong on other figures for
    the familiar reason. Re-expressing it as a fraction of the span is *also*
    wrong, because the right window depends on how far the walk travels, not on
    how tall the total is.

    And carrying over the reference's *proportions* fails too, which is the part
    worth writing down. Its C05X bridge is a small movement near the top of a
    tall scale - 24 kEUR on a base of 174, so it needs a lot of room beneath it
    for context, and the layout gives it room equal to 1.4 times the travel.
    Progressive's bridge travels 7,967 on a base of 3,489: more than twice its
    own starting height, because the company earns roughly three times its
    target margin. The same 1.4 multiple puts the bottom of the window at
    -7,804 - a negative axis under a measure that never goes negative.

    So the window is simply the walk plus a modest margin either side, held off
    a negative floor where every level is positive. It is not the reference's
    framing; the reference's framing describes a shape this data does not have.
    """
    low, high = min(levels), max(levels)
    travel = (high - low) or abs(high) or 1.0
    lo, hi = low - pad * travel, high + pad * travel
    if low >= 0:
        lo = max(lo, 0.0)
    return lo / span, hi / span


def _variance_bounds(values, pad: float = 0.1) -> tuple[float, float]:
    """A variance tier's axis, as fractions of its own span.

    The same lesson _walk_bounds records, in the other direction. The
    reference's window is asymmetric - far more room below zero than above -
    because its figures are mostly declines. Carried across to a book where
    every state grew, that window puts the ceiling at 0.63 of a span whose
    tallest bar is 1.0 by definition, and four bars are drawn outside their own
    axis.

    Zero is always an edge or inside, never padded away: a variance chart whose
    zero line is off the page has lost the only reference its readers have.
    """
    span = max(abs(v) for v in values)
    low, high = min(values) / span, max(values) / span
    return (low - pad if low < 0 else 0.0,
            high + pad if high > 0 else 0.0)


# A bar panel occupies a whole spreadsheet column, and Excel sizes that column
# from the panel's pixel width. Widest and narrowest are therefore both real
# costs: too wide and the column swallows the page, too narrow and its own
# heading will not fit.
_PANEL_MARGIN_PX = 4.0
_PANEL_MAX_CHARS = 18.0
_PANEL_MIN_CHARS = 8.0


def _bar_panels(template: D.Template, reference: dict,
                panel_value_width: float) -> dict:
    """Bar geometry for a table's panels, sized from the bars it will draw.

    Two separate faults in inheriting the reference's geometry, and this fixes
    both.

    **The ruler.** ``PanelGeometry.scale`` is pixels per kEUR, measured against
    a statement whose largest variance is a few hundred. Progressive's are in
    the thousands of millions, so the same ruler draws a bar a fraction of a
    pixel long - the flat-bars defect in its other direction: not clipped, but
    invisible.

    **The width, and where zero sits in it.** The reference puts zero about
    two-thirds of the way across each panel, because its variances run both
    ways and the left of the panel is where the adverse ones go. Progressive's
    are almost all favourable, so that two-thirds is blank - which is what made
    these columns look far too wide. Zero now goes where the data actually
    needs it, and the panel is only as wide as the bars either side of it.

    Panels sharing a scale group keep sharing one ruler, which is the whole
    claim of the notation: a month's miss and the year's are drawn at the same
    pixels per dollar, so they can be compared without arithmetic. Only the
    *widths* differ, and they differ because the bars do.
    """
    import ibcs_layout as _L

    budget = _PANEL_MAX_CHARS * 64.0 / panel_value_width
    floor = _PANEL_MIN_CHARS * 64.0 / panel_value_width

    # What each panel needs either side of zero, in data units.
    extents: dict[str, tuple[float, float]] = {}
    for key in reference:
        drawn = [v for v in template.tier(key).merged() if v is not None]
        extents[key] = (max(0.0, -min(drawn)), max(0.0, max(drawn)))

    # One ruler per group, set by whichever panel in it reaches furthest.
    rulers: dict[str, float] = {}
    for key, panel in reference.items():
        span = sum(extents[key]) or 1.0
        rulers[panel.scale_group] = min(
            rulers.get(panel.scale_group, float("inf")),
            (budget - 2.0 * _PANEL_MARGIN_PX) / span)

    panels = {}
    for key, panel in reference.items():
        scale = rulers[panel.scale_group]
        negative, positive = extents[key]
        width = max(floor,
                    2.0 * _PANEL_MARGIN_PX + (negative + positive) * scale)
        panels[key] = _L.PanelGeometry(
            round(width), round(_PANEL_MARGIN_PX + negative * scale),
            scale, scale_group=panel.scale_group)
    return panels


def _shared_unit_bounds(template_id: str, tallest: float, levels,
                        span: float, pad: float = 0.15,
                        lowest: float = 0.0) -> dict:
    """Bounds for a measure tier and the bridge that shares its ruler.

    ``same_unit_tiers`` means exactly what it says, and it is checked: the two
    tiers must show the same number of dollars per point, so their *ranges* are
    locked to the ratio of their plot heights. Only one of the two can be
    chosen freely; the other follows. Choosing both independently - which is
    what picking a sensible window for each looks like - silently puts the
    bridge on a different ruler from the columns it bridges, and the whole
    point of the sheet is that those two are comparable.

    So each tier states what it needs on its own, the larger need wins, and the
    other is derived from it:

    * the measure tier needs to reach its tallest column from zero;
    * the bridge needs to hold its walk with a margin either side.

    On the reference these are close, because its bridge travels 14% of its own
    base. Progressive earns roughly three times its target margin, so C05X's
    bridge travels more than twice its base - and there the bridge is the
    binding constraint, which is why that sheet's measure axis carries more
    headroom than the published render does. That is the arithmetic of a shared
    ruler, not a choice.
    """
    specs = {spec.key: spec for spec in L.LAYOUTS[template_id].tiers}
    ratio = specs["wf"].plot_extent / specs["measure"].plot_extent
    travel = max(levels) - min(levels)

    # A little headroom above the tallest column: an axis that ends exactly on
    # it draws the bar into the frame, which reads as a clipped bar and trips
    # the containment guard on a rounding hair. And a floor below zero only
    # where the data actually goes there - one month of 2025 was an
    # underwriting loss, and a zero floor drew it outside its own axis.
    floor = min(0.0, lowest * 1.02) / span
    measure_range = max(tallest * 1.02 / span - floor,
                        travel * (1.0 + 2.0 * pad) / span / ratio)
    wf_range = measure_range * ratio

    middle = (max(levels) + min(levels)) / 2.0 / span
    low = middle - wf_range / 2.0
    if min(levels) >= 0.0:
        low = max(low, 0.0)
    return {"measure": (floor, floor + measure_range),
            "wf": (low, low + wf_range)}


C06F_YEAR, C06F_PRIOR = 2025, 2024
C06F_CATCH_ALL = "all other"

# The reference's own span cells, read off the built recreation. Bounds
# declared as these fractions keep the proportions the published render has.
C06F_REFERENCE_SPAN = {"unit": 2145.3917313210045, "rel": 63.0}


def _c06f_states() -> list[tuple[str, float, float]]:
    """Each state, this year and last, largest contribution first."""
    table: dict[int, dict[str, float]] = {}
    for row in _rows("states.csv"):
        table.setdefault(int(row["year"]), {})[row["state"]] = float(row["npw"])
    now, prior = table[C06F_YEAR], table[C06F_PRIOR]

    named = [(name, now[name], prior[name]) for name in now
             if name not in ("total", C06F_CATCH_ALL) and name in prior]
    named.sort(key=lambda item: item[1] - item[2], reverse=True)
    catch_all = (C06F_CATCH_ALL, now[C06F_CATCH_ALL], prior[C06F_CATCH_ALL])
    return named + [catch_all]


C06F_ROWS_DATA = _c06f_states()
C06F_LABELS = tuple(
    "All other" if name == C06F_CATCH_ALL else name.title()
    for name, _ac, _py in C06F_ROWS_DATA)
C06F_AC = tuple(ac for _n, ac, _py in C06F_ROWS_DATA)
C06F_PY = tuple(py for _n, _ac, py in C06F_ROWS_DATA)
C06F_VAR_ABS = tuple(a - p for a, p in zip(C06F_AC, C06F_PY))
C06F_VAR_REL = tuple(v / p * 100.0 for v, p in zip(C06F_VAR_ABS, C06F_PY))

_c06f_total_ac = sum(C06F_AC)
_c06f_total_py = sum(C06F_PY)
_c06f_growth = _c06f_total_ac - _c06f_total_py
_c06f_named_growth = sum(C06F_VAR_ABS[:-1])
_c06f_leader = C06F_LABELS[0]
# The span the sheet will compute: the largest single cell any plotted column
# holds, which is the tallest summary stack rather than the tallest state.
_C06F_SPAN = max(_c06f_total_ac, _c06f_total_py, max(C06F_AC), max(C06F_PY))

C06F = D.Template(
    id="C06",
    variant="F",
    kind="bar chart with vertical waterfall",
    orientation="horizontal",
    title=D.TitleBlock(
        entity=ENTITY,
        measure="Net premiums written",
        unit="$m",
        period=str(C06F_YEAR),
        message=(
            f"Premium grew {_c06f_growth:+,.0f} $m "
            f"({_c06f_growth / _c06f_total_py * 100:+.1f}%) and every one of "
            f"the ten largest states grew. "
            f"But {C06F_VAR_ABS[-1] / _c06f_growth * 100:.0f}% of the "
            f"increase came from outside the top ten - "
            f"{C06F_VAR_ABS[-1]:+,.0f} $m against {_c06f_named_growth:+,.0f} "
            f"from all ten named states put together, {_c06f_leader} included"
        ),
    ),
    categories=C06F_LABELS,
    category_scenarios=("AC",) * len(C06F_LABELS),
    rows=tuple(D.Row(label=label, sign=1, kind="element")
               for label in C06F_LABELS),
    tiers=(
        D.Tier(key="measure", label="", kind="measure", printed=None,
               number_format="{:,.0f}",
               series=(D.Series("AC", list(C06F_AC)),
                       D.Series("PY", list(C06F_PY)))),
        D.Tier(key="wf", label="ΔPY", kind="waterfall", reference="PY",
               printed=None, number_format="{:+,.0f}",
               series=(D.Series("AC", list(C06F_VAR_ABS)),)),
        D.Tier(key="var_rel", label="ΔPY%", kind="variance_rel",
               reference="PY", printed=None, number_format="{:+.1f}",
               series=(D.Series("AC", list(C06F_VAR_REL)),)),
    ),
    # No plan row. The reference bridges prior year to actual and shows the
    # plan alongside; Progressive publishes no premium plan at all, still less
    # one by state, and the 96 combined ratio target says nothing about how
    # much premium to write. A summary bar with no basis would be the only
    # figure in the workbook that came from nowhere, so it is left out and the
    # notes say why.
    summary_rows=(
        D.Summary(label="PY", stack=(("PY", D.excel_round(_c06f_total_py)),),
                  before=True),
        D.Summary(label="AC", stack=(("AC", D.excel_round(_c06f_total_ac)),),
                  variance_abs=D.excel_round(_c06f_growth),
                  variance_rel=D.excel_round(
                      _c06f_growth / _c06f_total_py * 100.0, 1),
                  reference="PY"),
        D.Summary(label="ΔPY", stack=(),
                  variance_abs=D.excel_round(_c06f_growth),
                  variance_rel=D.excel_round(
                      _c06f_growth / _c06f_total_py * 100.0, 1),
                  reference="PY", span=("PY", "AC")),
    ),
    tier_bounds={
        **_shared_unit_bounds(
            "C06F", max(_c06f_total_ac, _c06f_total_py, max(C06F_AC)),
            [_c06f_total_py, _c06f_total_ac], _C06F_SPAN,
            lowest=min(min(C06F_AC), min(C06F_PY))),
        "var_rel": _variance_bounds(C06F_VAR_REL),
    },
    provenance=(
        D.SourceNote("AC", "filed",
                     f"Net premiums written by state, {C06F_YEAR} 10-K "
                     f"(Exhibit 13). The ten largest states are named and the "
                     f"rest are the filing's own All other line."),
        D.SourceNote("PY", "filed",
                     f"The same table for {C06F_PRIOR}. Both years tie to the "
                     f"annual premium built from twelve monthly releases - two "
                     f"documents, months apart."),
    ),
    source_ref="SEC EDGAR CIK 0000080661, 10-K Exhibit 13 state tables",
    notes=(
        "Both years are reported; the variance and its percentage are derived "
        "from them. The reference transcribes the actual and the percentage "
        "and derives the prior year - here there is no need.",
        "Rows are sorted by ΔPY descending. That ordering is what makes this "
        "variant F rather than E, G or H, so it is checked rather than "
        "assumed.",
        "'All other' is the filing's own catch-all for the forty states "
        "outside the top ten. It is pinned last whatever its size, because "
        "sorting the largest contributor into first place would present forty "
        "states as though they were one.",
        "No plan row: Progressive publishes no premium plan, by state or "
        "otherwise.",
    ),
)


# --------------------------------------------------------------------------- #
# C05X - the same months C07C draws, bridged to target instead of accumulated   #
# --------------------------------------------------------------------------- #
#
# C07C answers "is the year on track" with two cumulative lines. This answers
# "which months put it there" with a bridge: twelve steps from the target to
# the result. Same measure, same year, same figures - deliberately, because two
# notations over one set of numbers is the comparison the workbook exists to
# support, and a tie-out asserts the two sheets carry the identical series.
#
# The reference bridges to plan and shows prior year beside it, so this needs
# both: the 96 combined ratio target supplies the plan, and 2025's own result
# supplies the prior year.

C05X_YEAR = C07C_YEAR
C05X_PRIOR = C07C_YEAR - 1
C05X_SCENARIOS = C07C_SCENARIOS
C05X_LAST_ACTUAL = C07C_LAST_ACTUAL

C05X_MEASURE = C07C_VALUES["monthly"]
C05X_PLAN = C07C_VALUES["plan"]
# Prior-year months on the same basis, which is what the bridge walks from.
C05X_PY_VALUES = tuple(
    D.excel_round(_c07c_result((C05X_PRIOR, m))) for m in range(1, 13))
C05X_VAR_ABS = tuple(a - p for a, p in zip(C05X_MEASURE, C05X_PY_VALUES))
C05X_VAR_REL = tuple(v / p * 100.0
                     for v, p in zip(C05X_VAR_ABS, C05X_PY_VALUES))

# Last year's result, on the same basis, for the opening column.
C05X_PY_TOTAL = D.excel_round(sum(C05X_PY_VALUES))
_c05x_ac = D.excel_round(sum(C05X_MEASURE[:C05X_LAST_ACTUAL]))
_c05x_fc = D.excel_round(sum(C05X_MEASURE[C05X_LAST_ACTUAL:]))
_c05x_total = _c05x_ac + _c05x_fc
_c05x_plan_total = D.excel_round(sum(C05X_PLAN))
_c05x_vs_plan = _c05x_total - _c05x_plan_total
_c05x_vs_py = _c05x_total - C05X_PY_TOTAL
_c05x_best = max(range(12), key=lambda i: abs(C05X_VAR_ABS[i]))
_c05x_up = sum(1 for v in C05X_VAR_ABS if v > 0)

# The reference's own span cells, read off the built recreation.
C05X_REFERENCE_SPAN = {"unit": 174.0, "rel": 73.33333333333333}
# The span the sheet will compute. The bridge ends higher than any single
# measure column, so the walk's own top is what sets it.
_C05X_SPAN = max(C05X_PY_TOTAL, _c05x_total, _c05x_plan_total, _c05x_ac,
                 _c05x_fc, max(C05X_MEASURE))


def _c05x_split(values, keep):
    return [v if s == keep else None for v, s in zip(values, C05X_SCENARIOS)]


C05X = D.Template(
    id="C05",
    variant="X",
    kind="column chart with horizontal waterfall",
    title=D.TitleBlock(
        entity=ENTITY,
        measure="Underwriting result",
        unit="$m",
        period=str(C05X_YEAR),
        message=(
            f"The underwriting result is running {_c05x_vs_py:+,.0f} $m "
            f"({_c05x_vs_py / C05X_PY_TOTAL * 100:+.1f}%) ahead of "
            f"{C05X_PRIOR}, with {_c05x_up} of twelve months up and "
            f"{D.MONTHS[_c05x_best]} contributing "
            f"{C05X_VAR_ABS[_c05x_best]:+,.0f} on its own. Both years sit far "
            f"above the {TARGET_COMBINED_RATIO:.0f} combined ratio target, "
            f"drawn as the second opening column"
        ),
    ),
    categories=D.MONTHS,
    category_scenarios=C05X_SCENARIOS,
    rows=tuple(D.Row(label=month, sign=1, kind="element") for month in D.MONTHS),
    tiers=(
        D.Tier(key="var_rel", label="ΔPY%", kind="variance_rel",
               reference="PY", printed=None, number_format="{:+.1f}",
               series=(D.Series("AC", list(C05X_VAR_REL)),)),
        D.Tier(key="wf", label="ΔPY", kind="waterfall", reference="PY",
               printed=None, number_format="{:+,.0f}",
               series=(D.Series("AC", list(C05X_VAR_ABS)),)),
        # Three series per month, not four. The prior year is what the
        # bridge measures against, and the sheet already carries it: the
        # reference column is derived per month as measure less the variance,
        # and the PY *bar* belongs to the opening summary column alone - which
        # is exactly what the reference does. Adding a fourth monthly series
        # put a fourth bar in every cluster and made all of them narrower.
        D.Tier(key="measure", label="", kind="measure", printed=None,
               number_format="{:,.0f}",
               series=(D.Series("PL", list(C05X_PLAN), label="PL"),
                       D.Series("AC", _c05x_split(C05X_MEASURE, "AC"),
                                label="AC"),
                       D.Series("FC", _c05x_split(C05X_MEASURE, "FC"),
                                label="FC"))),
    ),
    # Plan first, prior year second: the bridge walks from the column
    # immediately above the months, and here that is prior year.
    summary_rows=(
        D.Summary(label=f"{C05X_YEAR} PL", stack=(("PL", _c05x_plan_total),),
                  before=True),
        D.Summary(label=f"{C05X_PRIOR} AC", stack=(("PY", C05X_PY_TOTAL),),
                  before=True),
        D.Summary(label=f"{C05X_YEAR} AC+FC",
                  stack=(("AC", _c05x_ac), ("FC", _c05x_fc)),
                  variance_abs=D.excel_round(_c05x_vs_py),
                  variance_rel=D.excel_round(
                      _c05x_vs_py / C05X_PY_TOTAL * 100.0, 1),
                  reference="PY"),
        D.Summary(label="vs PY", stack=(),
                  variance_abs=D.excel_round(_c05x_vs_py), reference="PY",
                  span=(f"{C05X_PRIOR} AC", f"{C05X_YEAR} AC+FC")),
        D.Summary(label="vs PL", stack=(),
                  variance_abs=D.excel_round(_c05x_vs_plan), reference="PL",
                  span=(f"{C05X_YEAR} PL", f"{C05X_YEAR} AC+FC")),
    ),
    tier_bounds={
        **_shared_unit_bounds(
            "C05X", max(C05X_PY_TOTAL, _c05x_total, max(C05X_MEASURE)),
            [C05X_PY_TOTAL, _c05x_total], _C05X_SPAN,
            lowest=min(min(C05X_MEASURE), min(C05X_PY_VALUES),
                       min(C05X_PLAN))),
        "var_rel": _variance_bounds(C05X_VAR_REL),
    },
    provenance=(
        D.SourceNote("AC", "derived",
                     "Earned premium times (100 - combined ratio), both from "
                     "the monthly 8-K exhibit. The same series C07C draws."),
        D.SourceNote("PY", "derived",
                     f"The same calculation over the twelve months of "
                     f"{C05X_PRIOR}."),
        D.SourceNote("PL", "assumed",
                     f"The {TARGET_COMBINED_RATIO:.0f} combined ratio target "
                     f"applied to the premium each month actually earned. The "
                     f"target is filed; using it as a monthly plan is this "
                     f"workbook's construction, and it is a floor the company "
                     f"aims to beat rather than a forecast."),
        D.SourceNote("FC", "assumed",
                     f"Remaining months are the prior year's at the "
                     f"year-to-date pace, earning the year-to-date margin - "
                     f"the basis C07C states and the same numbers."),
    ),
    source_ref="SEC EDGAR CIK 0000080661, monthly 8-K exhibits",
    notes=(
        "The same twelve months C07C draws, in the other notation: C07C "
        "accumulates them into two lines, this bridges them from target to "
        "result. A tie-out asserts the two sheets carry identical figures, so "
        "they cannot drift apart.",
        "Every step is an increase, because Progressive has beaten its own "
        "target in every month of the year. A bridge that only climbs is the "
        "honest picture here, not a missing colour.",
        "September is where measured becomes expected; the forecast steps are "
        "hatched and shaded as constructed.",
    ),
)


# --------------------------------------------------------------------------- #
# T01B - the reporting table: segment result, month and year to date            #
# --------------------------------------------------------------------------- #
#
# This is the table Progressive's own monthly release prints, in IBCS notation:
# each segment for the month and for the year so far, against last year and
# against target. The hierarchy is the company's, not a construction - Agency,
# Direct and Property are Personal Lines by the filing's own arithmetic, and
# Personal Lines plus Commercial plus the run-off businesses are companywide.
#
# Tier keys are period-neutral (``m_month``, ``m_ytd``) where the reference
# spells the month into them (``m_november``). Keeping the reference's keys
# would leave ``m_november`` holding July, which is the same trap as naming a
# 2011 panel "cologne".

T01B_YEAR, T01B_MONTH = C07C_YEAR, C07C_LAST_ACTUAL

# Label, role, and how the row behaves. ``None`` for a role means the row is
# computed from the ones above it.
T01B_STRUCTURE = (
    ("Agency", "agency", "element", None),
    ("Direct", "direct", "element", None),
    ("Property", "property", "element", None),
    ("Personal Lines", "personal_total", "subtotal", 3),
    ("Commercial Lines", "commercial", "element", None),
    ("Other", None, "element", None),
    ("Companywide", "companywide", "subtotal", "zero"),
)


def _t01b_block(year: int, month: int, ytd: bool) -> dict[str, tuple]:
    """One period's earned premium and underwriting result, by segment.

    The run-off businesses Progressive footnotes but never columns are the
    difference between companywide and the segments it does column, so "Other"
    is that remainder rather than an omission. C01A found the same gap on the
    premium line and showed it for the same reason: a table whose parts are
    supposed to make the whole cannot quietly drop the part that does not fit.
    """
    if ytd:
        source = {r["role"]: r for r in _rows("ytd_segments.csv")
                  if int(r["year"]) == year and int(r["through_month"]) == month}
    else:
        source = {r["role"]: r for r in _rows("monthly_segments.csv")
                  if int(r["year"]) == year and int(r["month"]) == month}

    earned, result = {}, {}
    for role, row in source.items():
        premium, ratio = _num(row.get("npe")), _num(row.get("combined_ratio"))
        if premium is None or ratio is None:
            continue
        earned[role] = premium
        result[role] = premium * (100.0 - ratio) / 100.0

    missing = [role for _l, role, _k, _s in T01B_STRUCTURE
               if role and role not in earned]
    if missing:
        raise ValueError(
            f"T01B: {year}-{month:02d} carries no segment result for "
            f"{missing} - the release for that month reports no ratio "
            f"against them")

    for store in (earned, result):
        store["other"] = (store["companywide"] - store["personal_total"]
                          - store["commercial"])
    return {"earned": earned, "result": result}


def _t01b_close(values: list) -> None:
    """Force the hierarchy shut: elements reported, subtotals derived."""
    agency, direct, prop, _personal, commercial, _other, companywide = values
    values[3] = D.excel_round(agency + direct + prop)          # Personal Lines
    values[5] = D.excel_round(                                  # Other
        companywide - agency - direct - prop - commercial)
    values[6] = D.excel_round(                                  # Companywide
        agency + direct + prop + commercial + values[5])


def _t01b_series(ytd: bool) -> dict:
    """One block of the table: actual, prior year and plan, row by row."""
    now = _t01b_block(T01B_YEAR, T01B_MONTH, ytd)
    prior = _t01b_block(T01B_YEAR - 1, T01B_MONTH, ytd)

    def column(block, key="result"):
        return [D.excel_round(block[key][role if role else "other"])
                for _label, role, _kind, _spans in T01B_STRUCTURE]

    actual, py = column(now), column(prior)
    # Make the table add up, which is the one thing a table has to do.
    #
    # Progressive reports each segment's combined ratio to a tenth, so results
    # derived from them do not sum to the derived total: Agency, Direct and
    # Property come to 5,797 where the reported Personal Lines ratio implies
    # 5,812. Fifteen million on six billion is the rounding, not a difference
    # of opinion - but a subtotal that does not equal the rows above it is
    # indefensible whatever its cause.
    #
    # So the subtotals are *derived* from the elements, and "Other" carries
    # whatever is left between the columned segments and the companywide figure
    # the release prints - the run-off businesses and the ratio rounding
    # together. Every other row is reported.
    for values in (actual, py):
        _t01b_close(values)
    # Plan is the workbook's one basis: the target margin on the premium each
    # segment actually earned, so the gap is margin against target and never a
    # premium miss - the same construction C05X and C07C use.
    plan = [D.excel_round(premium * TARGET_MARGIN / 100.0)
            for premium in column(now, "earned")]
    _t01b_close(plan)
    return {"AC": tuple(actual), "PY": tuple(py), "PL": tuple(plan),
            "earned": tuple(column(now, "earned"))}


T01B_MONTH_BLOCK = _t01b_series(ytd=False)
T01B_YTD_BLOCK = _t01b_series(ytd=True)


def _t01b_tiers() -> tuple:
    """Ten tiers: a measure and four variances, for the month and the year."""
    tiers = []
    period = f"{D.MONTHS[T01B_MONTH - 1]} {T01B_YEAR}"
    for suffix, block in (("month", T01B_MONTH_BLOCK), ("ytd", T01B_YTD_BLOCK)):
        actual = block["AC"]
        # The block name groups a measure column with the variances derived
        # from it; the leading underscore is how the renderer tells the
        # year-to-date block from the month's.
        name = period if suffix == "month" else f"_{period}"
        tiers.append(D.Tier(
            key=f"m_{suffix}", label="", kind="measure", printed=None,
            block=name, number_format="{:,.0f}",
            series=(D.Series("PY", list(block["PY"])),
                    D.Series("PL", list(block["PL"])),
                    D.Series("AC", list(actual)))))
        for scenario in ("PY", "PL"):
            reference = block[scenario]
            absolute = [a - r for a, r in zip(actual, reference)]
            tiers.append(D.Tier(
                key=f"d{scenario.lower()}_{suffix}", label=f"Δ{scenario}",
                kind="variance_abs", reference=scenario, printed=None,
                block=name, number_format="{:+,.0f}",
                series=(D.Series("AC", list(absolute)),)))
            tiers.append(D.Tier(
                key=f"d{scenario.lower()}p_{suffix}", label=f"Δ{scenario}%",
                kind="variance_rel", reference=scenario, printed=None,
                block=name, number_format="{:+.1f}",
                series=(D.Series("AC", [
                    None if r <= 0 else v / r * 100.0
                    for v, r in zip(absolute, reference)]),)))
    return tuple(tiers)


_t01b_labels = tuple(label for label, _r, _k, _s in T01B_STRUCTURE)
_t01b_total = T01B_YTD_BLOCK["AC"][-1]
_t01b_total_pl = T01B_YTD_BLOCK["PL"][-1]
# Only the columned segments - "Other" is a remainder, not a business, and
# ranking it beside the segments would compare a residual with a division.
_T01B_COLUMNED = [i for i, (_l, role, kind, _s) in enumerate(T01B_STRUCTURE)
                  if role and kind == "element"]
_t01b_best = max(_T01B_COLUMNED,
                 key=lambda i: (T01B_YTD_BLOCK["AC"][i]
                                - T01B_YTD_BLOCK["PL"][i]))

T01B = D.Template(
    id="T01",
    variant="B",
    kind="hierarchical table",
    title=D.TitleBlock(
        entity=ENTITY,
        measure="Underwriting result",
        unit="$m",
        period=f"{D.MONTHS[T01B_MONTH - 1]} {T01B_YEAR}",
        message=(
            f"Every columned segment beat the "
            f"{TARGET_COMBINED_RATIO:.0f} combined ratio target in the year "
            f"to date - companywide {_t01b_total:+,.0f} $m against the "
            f"{_t01b_total_pl:+,.0f} the target implies. "
            f"{_t01b_labels[_t01b_best]} contributes the largest excess "
            f"({T01B_YTD_BLOCK['AC'][_t01b_best] - T01B_YTD_BLOCK['PL'][_t01b_best]:+,.0f} $m) "
            f"on {100 - T01B_YTD_BLOCK['earned'][_t01b_best] / T01B_YTD_BLOCK['earned'][1] * 100:.0f}% "
            f"less earned premium than {_t01b_labels[1]}"
        ),
    ),
    categories=_t01b_labels,
    category_scenarios=("AC",) * len(_t01b_labels),
    rows=tuple(D.Row(label=label, sign=1, kind=kind, spans=spans)
               for label, _role, kind, spans in T01B_STRUCTURE),
    tiers=_t01b_tiers(),
    provenance=(
        D.SourceNote("AC", "derived",
                     "Segment earned premium and combined ratio, monthly 8-K "
                     "exhibit - the month block from the month's own table and "
                     "the year-to-date block from the year-to-date table on "
                     "the same page. The result is premium times "
                     "(100 - ratio)."),
        D.SourceNote("PY", "derived",
                     "The same two figures from the release twelve months "
                     "earlier, which is also the prior-year column that "
                     "release prints."),
        D.SourceNote("PL", "assumed",
                     f"The {TARGET_COMBINED_RATIO:.0f} combined ratio target "
                     f"on the premium each segment actually earned. The target "
                     f"is filed and company-wide; applying it to a single "
                     f"segment is this workbook's construction, and "
                     f"Progressive has never published a segment target."),
    ),
    source_ref="SEC EDGAR CIK 0000080661, monthly 8-K exhibits",
    notes=(
        "The hierarchy is Progressive's own: Agency, Direct and Property sum "
        "to Personal Lines in the filing's own arithmetic, and the five "
        "elements sum to companywide.",
        "'Other' is the run-off businesses the releases footnote but never "
        "column - the difference between companywide and the segments that "
        "are columned. Shown rather than absorbed, because a table whose "
        "parts are meant to make the whole cannot drop the part that does "
        "not fit.",
        "Applying a company-wide target to one segment is the weakest claim "
        "on the sheet, and it is the only one marked assumed.",
    ),
)


# --------------------------------------------------------------------------- #
# T03A and T04A - the underwriting statement, with and without bars             #
# --------------------------------------------------------------------------- #
#
# One statement, drawn twice: T03A in figures alone and T04A with the variance
# bars set among the columns. Both read the same block, so the two sheets cannot
# disagree - the same arrangement C05X and C07C have.
#
# An insurer's underwriting statement is short, and every line of it is either
# filed or forced. Premium less losses less the two expense lines *is* the
# underwriting result, and the three ratios beneath are those same figures over
# premium - so the ratio rows the template carries get a real workout rather
# than the single margin line the reference has.

T03A_YEAR, T03A_PRIOR = 2025, 2024

# (label, sign, kind, spans). Magnitudes are stored positive and the sign says
# which way the line runs, which is the convention the waterfall machinery and
# the statement renderer both already read.
T03A_LINES = (
    ("Net premiums earned", 1, "element", None, None),
    ("Losses and loss adjustment expense", -1, "element", None, None),
    ("Policy acquisition costs", -1, "element", None, None),
    ("Other underwriting expenses", -1, "element", None, None),
    ("Underwriting expenses", -1, "subtotal", 3, None),
    ("Underwriting result", 1, "subtotal", "zero", None),
    # A ratio row names the two rows it divides, because this statement has
    # more than one and they divide different lines. Indices are into this
    # tuple: 0 premium, 1 losses, 4 total underwriting expenses.
    #
    # The expense ratio is not a row: its numerator is two lines added
    # together, and a ratio row divides one row by another. It is the
    # difference between the two ratios that *are* rows, which is a subtraction
    # a reader can do on the page - and inventing a subtotal for
    # "everything except losses" to make the machinery fit would be arranging
    # the statement around the renderer.
    ("Loss and LAE ratio", 1, "ratio", None, (1, 0)),
    ("Combined ratio", 1, "ratio", None, (4, 0)),
)


def _t03a_statement(year: int) -> dict:
    """One year's underwriting statement, and the plan the target implies.

    Four figures are filed - earned premium, losses, acquisition costs and the
    combined ratio - and everything else is forced by them. Other underwriting
    expenses is what is left once the result the ratio implies is taken off, so
    the statement closes on premium exactly rather than to a tolerance.

    The plan holds the expense ratio where it actually landed and puts the
    whole of the target on losses, which is the same rule C08H uses: a combined
    ratio target says what may be spent in total, and expenses are the half a
    company controls directly.
    """
    earned = _X["premiums"][year]
    losses = _X["losses"][year]
    acquisition = _X["acquisition"][year]
    ratio = float(annual_segments()[year]["companywide"]["combined_ratio"])
    result = earned * (100.0 - ratio) / 100.0
    other = earned - losses - acquisition - result

    expense_ratio = (acquisition + other) / earned * 100.0
    plan_loss_ratio = TARGET_COMBINED_RATIO - expense_ratio
    plan_losses = earned * plan_loss_ratio / 100.0

    def close(premium, loss, acq, oth):
        """The four typed lines, then everything that follows from them.

        Rounded first, then divided. The workbook computes its ratios as
        formulas over the cells it prints, so a ratio derived here from
        unrounded components disagrees with the one the sheet shows - and the
        ratios are left unrounded because the number format is what does the
        displaying. The same lesson the return tree and the reserve
        roll-forward each taught once.
        """
        premium, loss, acq, oth = (D.excel_round(v)
                                   for v in (premium, loss, acq, oth))
        expenses = loss + acq + oth
        return [premium, loss, acq, oth, expenses, premium - expenses,
                loss / premium * 100.0, expenses / premium * 100.0]

    return {
        "AC": tuple(close(earned, losses, acquisition, other)),
        "PL": tuple(close(earned, plan_losses, acquisition, other)),
        # Not a row, but the message and the tie-outs both need it.
        "expense_ratio": expense_ratio,
    }


T03A_AC = _t03a_statement(T03A_YEAR)
T03A_PY = _t03a_statement(T03A_PRIOR)
T03A_VALUES = {"AC": T03A_AC["AC"], "PY": T03A_PY["AC"], "PL": T03A_AC["PL"]}


def _t03a_tiers(with_prior: bool) -> tuple:
    """The measure column and the variances beside it.

    T04A drops the prior-year pair: its bars are the plan variance, and a
    statement carrying four variance columns plus two panels has no room left
    for the figures.
    """
    actual = T03A_VALUES["AC"]
    tiers = [D.Tier(key="m", label="", kind="measure", printed=None,
                    number_format="{:,.0f}",
                    series=(D.Series("PY", list(T03A_VALUES["PY"])),
                            D.Series("PL", list(T03A_VALUES["PL"])),
                            D.Series("AC", list(actual))))]
    references = ("PY", "PL") if with_prior else ("PL",)
    for scenario in references:
        reference = T03A_VALUES[scenario]
        absolute = [a - r for a, r in zip(actual, reference)]
        tiers.append(D.Tier(
            key=f"d{scenario.lower()}", label=f"Δ{scenario}",
            kind="variance_abs", reference=scenario, printed=None,
            number_format="{:+,.0f}",
            series=(D.Series("AC", list(absolute)),)))
        tiers.append(D.Tier(
            key=f"d{scenario.lower()}p", label=f"Δ{scenario}%",
            kind="variance_rel", reference=scenario, printed=None,
            number_format="{:+.1f}",
            series=(D.Series("AC", [None if r <= 0 else v / r * 100.0
                                    for v, r in zip(absolute, reference)]),)))
    return tuple(tiers)


_t03a_labels = tuple(label for label, _s, _k, _sp, _r in T03A_LINES)
_t03a_rows = tuple(D.Row(label=label, sign=sign, kind=kind, spans=spans,
                         ratio_of=ratio)
                   for label, sign, kind, spans, ratio in T03A_LINES)
_t03a_result, _t03a_plan_result = T03A_VALUES["AC"][5], T03A_VALUES["PL"][5]
T03A_EXPENSE_RATIO = {k: T03A_AC["expense_ratio"] if k != "PY"
                      else T03A_PY["expense_ratio"] for k in ("AC", "PY", "PL")}
_t03a_combined = T03A_VALUES["AC"][7]


def _t03a_expense(scenario: str) -> float:
    """The expense ratio: what the combined ratio has left after losses."""
    return T03A_VALUES[scenario][7] - T03A_VALUES[scenario][6]


def _t03a_message() -> str:
    return (
        f"A {_t03a_combined:.1f} combined ratio against the "
        f"{TARGET_COMBINED_RATIO:.0f} target left an underwriting result of "
        f"{_t03a_result:+,.0f} $m, {_t03a_result - _t03a_plan_result:+,.0f} "
        f"more than the target implies. The gain over {T03A_PRIOR} is all in "
        f"losses: the loss ratio fell "
        f"{T03A_VALUES['PY'][6] - T03A_VALUES['AC'][6]:.1f} points while the "
        f"expense ratio rose "
        f"{_t03a_expense('AC') - _t03a_expense('PY'):.1f}"
    )


_T03A_PROVENANCE = (
    D.SourceNote("AC", "derived",
                 "Earned premium, losses and loss adjustment expense and "
                 "policy acquisition costs are tagged in the 10-K; the "
                 "combined ratio is the monthly release's full-year figure. "
                 "Other underwriting expenses and the result follow from "
                 "them, so the statement closes on premium exactly."),
    D.SourceNote("PY", "derived", f"The same statement for {T03A_PRIOR}."),
    D.SourceNote("PL", "assumed",
                 f"The {TARGET_COMBINED_RATIO:.0f} combined ratio target with "
                 f"the expense ratio held where it actually landed, so the "
                 f"whole of the target falls on losses. The target is filed; "
                 f"resolving it into a statement line is this workbook's "
                 f"construction."),
)

_T03A_NOTES = (
    "Four filed figures determine the statement: earned premium, losses, "
    "acquisition costs and the combined ratio. Other underwriting expenses is "
    "the remainder, so premium less every expense line is the result to the "
    "million.",
    "The three ratio rows are those same figures over earned premium, which "
    "is why they are drawn in percent and their absolute variances in "
    "percentage points.",
    "Plan holds the expense ratio at actual and puts the target's whole "
    "weight on losses - the rule C08H uses, so the workbook has one plan "
    "basis rather than two.",
)

T03A = D.Template(
    id="T03",
    variant="A",
    kind="statement table",
    title=D.TitleBlock(
        entity=ENTITY, measure="Underwriting result", unit="$m",
        period=str(T03A_YEAR), message=_t03a_message()),
    categories=_t03a_labels,
    category_scenarios=("AC",) * len(_t03a_labels),
    rows=_t03a_rows,
    tiers=_t03a_tiers(with_prior=True),
    provenance=_T03A_PROVENANCE,
    source_ref="SEC EDGAR CIK 0000080661, 10-K and monthly 8-K exhibits",
    notes=_T03A_NOTES,
)


T04A = D.Template(
    id="T04",
    variant="A",
    kind="statement table with integrated bars",
    title=D.TitleBlock(
        entity=ENTITY, measure="Underwriting result", unit="$m",
        period=str(T03A_YEAR), message=_t03a_message()),
    categories=_t03a_labels,
    category_scenarios=("AC",) * len(_t03a_labels),
    rows=_t03a_rows,
    tiers=_t03a_tiers(with_prior=False),
    panel_tiers=("dpl", "dplp"),
    provenance=_T03A_PROVENANCE,
    source_ref="SEC EDGAR CIK 0000080661, 10-K and monthly 8-K exhibits",
    notes=_T03A_NOTES + (
        "The same statement T03A draws, with the plan variance set among the "
        "columns as bars. Both read one block, so the two sheets cannot "
        "disagree.",
    ),
)

import ibcs_layout as _L04

T04A.panel_geometry.update(
    _bar_panels(T04A, _L04.T04A_PANELS, _L04.T04A_TABLE.panel_value_width))


# --------------------------------------------------------------------------- #
# T02A - the same table T01B draws, with the plan variance set among the bars   #
# --------------------------------------------------------------------------- #
#
# T01B in figures alone carries four variance columns; this one carries two and
# draws them, which is the trade the variant makes. Both read the same two
# blocks, so the sheets cannot disagree - the arrangement T03A and T04A have,
# and C05X and C07C before them.


def _t02a_tiers() -> tuple:
    """Six tiers: the measure and the plan variance, month and year to date.

    Prior year is dropped rather than shrunk. A bar column needs the width of
    several figure columns, and a table cannot carry both references and both
    panels; the reference makes the same choice.
    """
    return tuple(tier for tier in _t01b_tiers()
                 if not tier.key.startswith(("dpy", "dpyp")))


T02A = D.Template(
    id="T02",
    variant="A",
    kind="table with integrated bars",
    title=D.TitleBlock(
        entity=ENTITY,
        measure="Underwriting result",
        unit="$m",
        period=f"{D.MONTHS[T01B_MONTH - 1]} {T01B_YEAR}",
        message=(
            f"Against the {TARGET_COMBINED_RATIO:.0f} combined ratio target, "
            f"every columned segment is ahead for the year to date - "
            f"{_t01b_total:+,.0f} $m against {_t01b_total_pl:+,.0f}. The bars "
            f"are the gap to target, drawn at one scale for the month and the "
            f"year so the two can be read against each other"
        ),
    ),
    categories=_t01b_labels,
    category_scenarios=("AC",) * len(_t01b_labels),
    rows=tuple(D.Row(label=label, sign=1, kind=kind, spans=spans)
               for label, _role, kind, spans in T01B_STRUCTURE),
    tiers=_t02a_tiers(),
    # Which tiers are drawn as bars rather than printed as figures. Without
    # this the columns still exist, still hold the right numbers, and simply
    # show them - a table that looks finished and is not the template.
    panel_tiers=("dpl_month", "dplp_month", "dpl_ytd", "dplp_ytd"),
    provenance=T01B.provenance,
    source_ref="SEC EDGAR CIK 0000080661, monthly 8-K exhibits",
    notes=T01B.notes + (
        "The same two blocks T01B draws, with the plan variance as bars "
        "instead of two further figure columns. Both read one block, so the "
        "sheets cannot disagree.",
    ),
)

import ibcs_layout as _L02

# The reference's four panels, named by the keys this dataset actually uses.
# Only the scale groups are taken from them - every dimension is recomputed.
_T02A_REFERENCE_PANELS = {
    "dpl_month": _L02.PanelGeometry(112, 67, 0.0, scale_group="unit"),
    "dplp_month": _L02.PanelGeometry(224, 110, 0.0, scale_group="rel"),
    "dpl_ytd": _L02.PanelGeometry(224, 132, 0.0, scale_group="unit"),
    "dplp_ytd": _L02.PanelGeometry(224, 110, 0.0, scale_group="rel"),
}

T02A.panel_geometry.update(
    _bar_panels(T02A, _T02A_REFERENCE_PANELS,
                _L02.T02A_TABLE.panel_value_width))


TEMPLATES: dict[str, D.Template] = {"T02A": T02A, "T03A": T03A, "T04A": T04A, "T01B": T01B, "C05X": C05X, "C06F": C06F, "C07C": C07C, "C08H": C08H, "C11A": C11A, "C09C": C09C, "C04A": C04A, "C13D": C13D, "C01A": C01A, "C02A": C02A,
                                    "C03A": C03A, "C10D": C10D, "C12A": C12A}
