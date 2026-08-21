"""Tie-outs for the Progressive dataset.

**This is the real gate, and it exists because the usual one is not.**
``ibcs_data.check_ties`` takes a ``Template`` and then ignores it: every
per-template check reads module globals from ``ibcs_data``. Handed a template
built from other figures it re-verifies the Institute's numbers against
themselves, passes, and proves nothing about the data actually being drawn.
``build()`` gates on it, so without this the workbook would be built on an
assertion that cannot fail.

Every check here reads the template it was handed.
"""
from __future__ import annotations

import sys

import ibcs_data as D
import ibcs_data_alt as A
import ibcs_layout as L


class TieError(AssertionError):
    """A derived figure disagrees with what it was derived from."""


def run(template: D.Template, checks: list) -> int:
    results = []

    def check(name: str, ok: bool, detail: str) -> None:
        results.append(ok)
        print(f"  [{'PASS' if ok else 'FAIL'}] {name} - {detail}")

    for fn in checks:
        fn(template, check)
    failed = results.count(False)
    return failed


def _c03a(t: D.Template, check) -> None:
    py = t.tier("measure").series_for("PY").values
    measure = t.tier("measure").merged()
    var_abs = t.tier("var_abs").merged()
    var_rel = t.tier("var_rel").merged()

    check("every series is the full category length",
          all(len(x) == len(t.categories)
              for x in (py, measure, var_abs, var_rel)),
          f"{len(t.categories)} categories")

    worst = max(abs((m - p) - v)
                for m, p, v in zip(measure, py, var_abs))
    check("absolute variance is actual less prior year",
          worst < 1e-9, f"largest disagreement {worst:.2e}")

    worst = max(abs(v / p * 100.0 - r)
                for v, p, r in zip(var_abs, py, var_rel))
    check("relative variance is absolute over prior year",
          worst < 1e-9, f"largest disagreement {worst:.2e}")

    # No two scenarios may claim the same period. A slot filled twice draws two
    # bars in one place; a slot filled never leaves a gap that reads as zero.
    tier = t.tier("measure")
    ac = tier.series_for("AC").values
    fc = tier.series_for("FC").values
    clash = [i for i, (a, f) in enumerate(zip(ac, fc))
             if (a is not None) == (f is not None)]
    check("actual and forecast partition the year",
          not clash, f"overlapping or empty periods: {clash or 'none'}")

    # The summary must be the series it summarises, not a figure typed beside
    # it. This is the check that catches a forecast changed without its total.
    summary = t.summary
    stacked = sum(v for _s, v in summary.stack)
    total = sum(measure)
    check("the year total is the twelve months",
          abs(stacked - total) <= 1.0,
          f"summary {stacked:,.0f} vs months {total:,.0f}")

    check("the summary variance matches its own stack",
          abs(summary.variance_abs - (stacked - sum(py))) <= 1.0,
          f"{summary.variance_abs:+,.0f} vs {stacked - sum(py):+,.0f}")

    # The summary's own variance pair must be internally consistent: the
    # printed percentage, applied back to the total, must reproduce the printed
    # absolute.
    #
    # The tolerance is *derived*, not chosen. The percentage prints to one
    # decimal, so it stands for a band half a tenth wide, and half a tenth of a
    # total is worth more when the total is bigger. ibcs_data makes the same
    # check with a flat +-1.5, which is right at the Institute's kEUR
    # magnitudes and wrong here: at Progressive's scale one decimal place of
    # percent is worth about +-44, so a correct figure fails a fixed tolerance.
    # A tolerance that does not scale with the number it tolerates is not a
    # tolerance, it is a magnitude assumption.
    tol = stacked * 0.05 / 100.0
    implied = stacked - stacked / (1 + summary.variance_rel / 100.0)
    check("the summary percentage reproduces its own absolute",
          abs(implied - summary.variance_abs) <= tol,
          f"{summary.variance_rel:+.1f}% of {stacked:,.0f} implies "
          f"{implied:+,.1f}, printed {summary.variance_abs:+,.0f} "
          f"(tolerance +-{tol:,.1f} from one decimal place)")

    # The forecast must follow the stated rule, or the footnote is a lie.
    first_fc = t.category_scenarios.index("FC")
    worst = max(abs(measure[i] - D.excel_round(py[i] * A.C03A_PACE))
                for i in range(first_fc, len(measure)))
    check("the forecast follows the rule the sheet prints",
          worst < 1e-9,
          f"prior year grown {A.C03A_PACE * 100 - 100:+.2f}%, "
          f"largest disagreement {worst:.2e}")

    # Provenance: every scenario drawn must say where it came from, or the
    # shading and the Sources sheet have nothing to report.
    drawn = {s for s in t.category_scenarios} | {
        s.scenario for tier in t.tiers for s in tier.series}
    stated = {n.scenario for n in t.provenance}
    check("every scenario drawn declares its provenance",
          drawn <= stated, f"drawn={sorted(drawn)} stated={sorted(stated)}")


def _c10d(t: D.Template, check) -> None:
    pts = t.points
    check("every point carries all three measures",
          all(p.size is not None and p.x is not None and p.y is not None
              for p in pts),
          f"{len(pts)} points")

    check("every bubble has positive area",
          all(p.size > 0 for p in pts),
          f"smallest {min(p.size for p in pts):,.0f}")

    # The two axes are components of one ratio, so their sum is the combined
    # ratio the company reports. Checking it here is what stops a point being
    # plotted from one year's expense ratio and another year's losses.
    seg = A.annual_segments()
    worst = 0.0
    for p in pts:
        year = A.C10D_PY_YEAR if p.scenario == "PY" else A.C10D_AC_YEAR
        role = next(r for r, lab in A.C10D_LABELS.items() if lab == p.entity)
        row = seg[year][role]
        worst = max(worst, abs((p.x + p.y) - float(row["combined_ratio"])))
    check("expense plus loss is the reported combined ratio",
          worst <= 0.05,
          f"largest disagreement {worst:.3f} points")

    # The acquisition is the one unit drawn without a prior year. If it ever
    # gained one the notation would be claiming something untrue of it.
    acq = {p.entity for p in pts if p.scenario == "ACQ"}
    twinned = {p.entity for p in pts if p.scenario == "PY"}
    check("the acquisition has no prior-year twin",
          acq and not (acq & twinned),
          f"ACQ={sorted(acq)}, prior-year units={sorted(twinned)}")

    check("every other unit is drawn twice",
          all(sum(1 for p in pts if p.entity == e) == 2 for e in twinned),
          f"{len(twinned)} units carry a prior year and an actual")

    drawn = {p.scenario for p in pts}
    stated = {n.scenario for n in t.provenance}
    check("every scenario drawn declares its provenance",
          drawn <= stated, f"drawn={sorted(drawn)} stated={sorted(stated)}")


def _c02a(t: D.Template, check) -> None:
    panel = t.structure_panels[0]
    cats = list(panel.categories)
    segs = panel.segments

    check("every segment spans every category",
          all(len(s.values) == len(cats) for s in segs),
          f"{len(cats)} categories, {len(segs)} segments")

    # Each bar is the earned premium it decomposes. Rounding is allowed for one
    # unit per segment and no more: the parts are rounded independently, so a
    # bar of three can legitimately miss its own total by up to 1.5.
    seg = A.annual_segments()[A.C02A_YEAR]
    worst = 0.0
    for role, label in A.C02A_LABELS.items():
        i = cats.index(label)
        drawn = sum(s.values[i] for s in segs)
        worst = max(worst, abs(drawn - float(seg[role]["npe"])))
    check("each segment's parts are its earned premium",
          worst <= 1.5, f"largest disagreement {worst:.2f} $m")

    # Subtotals must be the bars above them, or the hierarchy is decoration.
    i_personal = cats.index("Personal Lines")
    i_company = cats.index("Companywide")
    parts = [cats.index(A.C02A_LABELS[r]) for r in A.C02A_PERSONAL]
    worst_p = worst_c = 0.0
    for s in segs:
        worst_p = max(worst_p,
                      abs(sum(s.values[i] for i in parts) - s.values[i_personal]))
        worst_c = max(worst_c,
                      abs(s.values[i_personal]
                          + s.values[cats.index("Commercial")]
                          - s.values[i_company]))
    check("Personal Lines is Agency plus Direct plus Property",
          worst_p <= 1.5, f"largest disagreement {worst_p:.2f} $m")
    check("Companywide is Personal Lines plus Commercial",
          worst_c <= 1.5, f"largest disagreement {worst_c:.2f} $m")

    check("no bar carries a negative part",
          all(v >= 0 for s in segs for v in s.values),
          "losses, expenses and result are all non-negative this year")

    drawn = set(panel.category_scenarios)
    stated = {n.scenario for n in t.provenance}
    check("every scenario drawn declares its provenance",
          drawn <= stated, f"drawn={sorted(drawn)} stated={sorted(stated)}")


def _c01a(t: D.Template, check) -> None:
    panels = {p.key: p for p in t.structure_panels}
    segment, state = panels["segment"], panels["state"]

    check("every segment spans every year",
          all(len(s.values) == len(segment.categories)
              for s in segment.segments),
          f"{len(segment.categories)} years, {len(segment.segments)} segments")

    # Each year's stack is that year's companywide premium. Four figures each
    # rounded to a million can miss a fifth by two, and no more.
    seg = A.annual_segments()
    worst, worst_year = 0.0, None
    for i, year in enumerate(segment.categories):
        drawn = sum(s.values[i] for s in segment.segments)
        want = float(seg[int(year)]["companywide"]["npw"])
        if abs(drawn - want) > worst:
            worst, worst_year = abs(drawn - want), year
    # Tight, because the residual is now a segment of its own rather than a
    # discrepancy. What is left is the rounding of the drawn figures.
    check("each year's stack is that year's companywide premium",
          worst <= 1.0, f"largest disagreement {worst:.1f} $m in {worst_year}")

    # The state panel must be the whole book, not just the states named.
    drawn = sum(s.values[0] for s in state.segments)
    want = A._STATES[A.C01A_STATE_YEAR]["total"]
    check("the state panel is the whole book",
          abs(drawn - want) <= 1.0,
          f"{drawn:,.0f} vs the 10-K's {want:,.0f} $m")

    # And the two panels must agree about the year they share - they come from
    # different documents, so this is the check that they are the same measure.
    shared = state.categories[0]
    if shared in segment.categories:
        i = list(segment.categories).index(shared)
        by_segment = sum(s.values[i] for s in segment.segments)
        check("the two panels agree about the year they share",
              abs(by_segment - drawn) <= 5.0,
              f"by segment {by_segment:,.0f} vs by state {drawn:,.0f} $m "
              f"- independent rounding in two filings")

    check("the largest state is not larger than the book",
          max(s.values[0] for s in state.segments) <= want,
          "no part exceeds its own total")

    stated = {n.scenario for n in t.provenance}
    check("every scenario drawn declares its provenance",
          set(segment.category_scenarios) | set(state.category_scenarios)
          <= stated, f"stated={sorted(stated)}")


def _c12a(t: D.Template, check) -> None:
    import ibcs_layout as LAY

    cats = list(t.categories)
    py = t.tier("wf_py").merged()
    ac = t.tier("wf_ac").merged()
    var_abs = t.tier("var_abs").merged()
    var_rel = t.tier("var_rel").merged()

    check("every series spans the statement",
          all(len(x) == len(cats) for x in (py, ac, var_abs, var_rel)),
          f"{len(cats)} lines")

    # The statement must walk: every subtotal is the signed lines above it.
    # This is the check that catches a line changed without its total.
    worst, worst_row = 0.0, None
    for values, label in ((py, "PY"), (ac, "AC")):
        for i, row in enumerate(t.rows):
            if row.kind != "subtotal":
                continue
            running = sum(r.sign * values[j]
                          for j, r in enumerate(t.rows[:i])
                          if r.kind != "subtotal")
            if abs(running - values[i]) > worst:
                worst, worst_row = abs(running - values[i]), f"{label} {cats[i]}"
    check("every subtotal is the lines above it",
          worst <= 1.0, f"largest disagreement {worst:.1f} $m at {worst_row}")

    # Variance carries impact, not sign: a bigger expense is adverse, so the
    # row's own sign is applied. Getting this backwards paints every cost line
    # the wrong colour and the chart still looks plausible.
    worst = max(abs(row.sign * (a - p) - v)
                for row, a, p, v in zip(t.rows, ac, py, var_abs))
    check("absolute variance carries the line's own sign",
          worst < 1e-9, f"largest disagreement {worst:.2e}")

    # The clipping notation is the point of this template, so the shipped file
    # has to exercise it rather than describe it.
    lo, hi = LAY.C12A_REL_LIMIT
    outliers = [i for i, v in enumerate(var_rel)
                if v is not None and (v > hi or v < lo)]
    check("the overflow marks are exercised by the data",
          len(outliers) >= 2,
          f"{len(outliers)} line(s) beyond {lo:g}..{hi:g}: "
          f"{', '.join(cats[i] for i in outliers)}")

    check("the outlier list is derived, not typed",
          tuple(outliers) == tuple(A.C12A_OUTLIERS),
          "C12A_OUTLIERS comes from the layout's own limit")

    # A percentage of a negative base is not meaningful, and the workbook
    # suppresses it. Anything this template still draws must have a positive
    # reference behind it.
    drawn_on_bad_base = [cats[i] for i, p in enumerate(py)
                         if p <= 0 and var_rel[i] is not None]
    check("no relative variance is drawn on a non-positive base",
          not drawn_on_bad_base,
          f"lines with a non-positive prior year: "
          f"{[cats[i] for i, p in enumerate(py) if p <= 0] or 'none'}; "
          f"drawn anyway: {drawn_on_bad_base or 'none'}")

    stated = {n.scenario for n in t.provenance}
    check("every scenario drawn declares its provenance",
          {"AC", "PY"} <= stated, f"stated={sorted(stated)}")


def _c13d(t: D.Template, check) -> None:
    import ibcs_layout as LAY

    keys = list(A.C13D_KEYS)
    check("fifteen panels of twelve months",
          len(keys) == 15 and all(len(A.C13D_MONTHLY[k]) == 12 for k in keys),
          f"{len(keys)} panels, {len(t.categories)} categories")

    # Every panel key must be drawn, or a year is silently missing from a grid
    # that looks complete.
    grid = t.panel_grids["uniform"]
    drawn = {c.key for c in grid.drawn()}
    check("every panel key is drawn exactly once",
          drawn == set(keys) and len(grid.drawn()) == len(keys),
          f"{len(grid.drawn())} cells drawn, {grid.rows}x{grid.cols} grid")

    # The reference really is the mean of the panels, month by month.
    worst = 0.0
    for m in range(12):
        mean = sum(A.C13D_MONTHLY[k][m] for k in keys) / len(keys)
        worst = max(worst, abs(mean - A.C13D_AVERAGE[m]))
    check("the reference series is the mean of the panels",
          worst <= 0.5, f"largest disagreement {worst:.2f} $m")

    # Each panel's variance reproduces from its own figures.
    worst = 0.0
    for k in keys:
        for m in range(12):
            total = sum(A.C13D_MONTHLY[j][m] for j in keys)
            want = (A.C13D_MONTHLY[k][m] * len(keys) - total) / total * 100.0
            worst = max(worst, abs(t.tier(k).merged()[m] - want))
    check("every variance reproduces from the panel figures",
          worst < 1e-9, f"largest disagreement {worst:.2e}")

    # Nothing may exceed what the panel layout can draw, or a pin runs off its
    # own panel and the grid stops sharing a scale.
    limit = LAY.panel_layout_for("C13D").maximum
    biggest = max(abs(v) for k in keys for v in t.tier(k).merged())
    check("no panel exceeds the layout's maximum",
          biggest <= limit, f"largest |variance| {biggest:.0f}% against {limit:g}")

    # The callout must point at the bar the message names.
    key, month = t.annotations[0].target
    check("the callout points at the panel the message names",
          t.tier(key).merged()[month] == max(
              v for k in keys for v in t.tier(k).merged()),
          f"{key} {D.MONTHS[month]} at "
          f"{t.tier(key).merged()[month]:+,.0f}%")

    # And each year's twelve months must still be that year's annual total.
    seg = {int(r["year"]): r for r in A._rows("monthly_headline.csv")}
    stated = {n.scenario for n in t.provenance}
    check("every scenario drawn declares its provenance",
          set(t.category_scenarios) <= stated, f"stated={sorted(stated)}")


def _c04a(t: D.Template, check) -> None:
    measure = t.tier("measure")
    ac = measure.series_for("AC").values
    pl = measure.series_for("PL").values
    va = t.tier("var_abs").merged()
    vr = t.tier("var_rel").merged()

    check("every series spans the year",
          all(len(x) == 12 for x in (ac, pl, va, vr)),
          f"{len(t.categories)} categories")

    # The ordering *is* variant A. If it stops being sorted the template is a
    # different template.
    check("months are ranked by the gap to plan, largest first",
          all(a >= b - 1e-9 for a, b in zip(va, va[1:])),
          f"{va[0]:+,.0f} down to {va[-1]:+,.0f} $m")

    worst = max(abs((a - p) - v) for a, p, v in zip(ac, pl, va))
    check("absolute variance is actual less plan",
          worst < 1e-9, f"largest disagreement {worst:.2e}")

    worst = max(abs(v / p * 100.0 - r) for v, p, r in zip(va, pl, vr))
    check("relative variance is absolute over plan",
          worst < 1e-9, f"largest disagreement {worst:.2e}")

    # The plan must follow the basis the sheet prints, or the footnote is a lie.
    worst = 0.0
    for cat, p in zip(t.categories, pl):
        month = D.MONTHS.index(cat) + 1
        npe = float(A._M[(A.C04A_YEAR, month)]["npe"])
        worst = max(worst, abs(p - D.excel_round(npe * A.TARGET_MARGIN / 100.0)))
    check("the plan is the stated 96 combined ratio basis",
          worst < 1e-9,
          f"earned premium x {A.TARGET_MARGIN:.0f}%, largest disagreement "
          f"{worst:.2e}")

    # Both colours must appear, or the variance notation is unexercised.
    check("the year carries both favourable and adverse months",
          any(v > 0 for v in va) and any(v < 0 for v in va),
          f"{sum(1 for v in va if v > 0)} above plan, "
          f"{sum(1 for v in va if v < 0)} below")

    # And the constructed scenario must say so, or nothing gets shaded.
    bases = {n.scenario: n.basis for n in t.provenance}
    check("the plan is declared as assumed, the actual as filed",
          bases.get("PL") == "assumed" and bases.get("AC") == "filed",
          f"{bases}")


def _c09c(t: D.Template, check) -> None:
    pts = t.points
    iso = t.iso_curves

    check("every point carries both coordinates",
          all(p.x is not None and p.y is not None for p in pts),
          f"{len(pts)} points")

    # The iso-curves are only meaningful in the positive quadrant, and a point
    # outside it would not be drawn at all.
    check("every point is inside the quadrant the curves live in",
          all(p.x > 0 and p.y > 0 for p in pts),
          f"margin {min(p.x for p in pts):.1f}..{max(p.x for p in pts):.1f}%, "
          f"premium {min(p.y for p in pts):,.0f}..{max(p.y for p in pts):,.0f}")

    # Each group must be a complete run of the same months, or a cluster is
    # comparing different periods with itself.
    counts = {g.name: sum(1 for p in pts if p.group == g.name) for g in t.groups}
    check("every segment contributes the same months",
          len(set(counts.values())) == 1,
          f"{counts}")

    # The claim the message makes, recounted from the points it is made about.
    counted = sum(1 for p in pts
                  if p.group == iso.group
                  and D.c09c_gross_profit(p) >= iso.segment)
    check("the message's count is the count on the page",
          counted == iso.count,
          f"{iso.group} at {iso.segment:,.0f} $m or more: {counted} points, "
          f"message says {iso.count}")

    # The third measure is the product of the two axes - that is what makes the
    # curves curves rather than decoration.
    worst = max(abs(D.c09c_gross_profit(p) - p.x * p.y / 100.0) for p in pts)
    check("the derived measure is the product of the axes",
          worst < 1e-9, f"largest disagreement {worst:.2e}")

    check("the shaded level is the outermost curve",
          iso.segment == max(iso.levels),
          f"levels {iso.levels}, segment {iso.segment:,.0f}")

    stated = {n.scenario for n in t.provenance}
    check("every scenario drawn declares its provenance",
          {p.scenario for p in pts} <= stated, f"stated={sorted(stated)}")


def _c11a(t: D.Template, check) -> None:
    v = A.C11A_VALUES
    n = len(t.categories)
    check("every node spans every year",
          all(len(v[k]) == n for k in v), f"{n} years, {len(v)} nodes")

    # The links are the tree. If they stop holding, the connectors draw an
    # arithmetic the boxes are not doing.
    worst = max(abs(r / s * 100.0 - m)
                for r, s, m in zip(v["return"], v["net_sales"], v["ros"]))
    check("margin is result over premium",
          worst < 1e-9, f"largest disagreement {worst:.2e}")
    worst = max(abs(s / c - tn)
                for s, c, tn in zip(v["net_sales"], v["capital"], v["turnover"]))
    check("turnover is premium over equity",
          worst < 1e-9, f"largest disagreement {worst:.2e}")
    worst = max(abs(m * tn - r)
                for m, tn, r in zip(v["ros"], v["turnover"], v["roi"]))
    check("return on equity is margin times turnover",
          worst < 1e-9, f"largest disagreement {worst:.2e}")

    # One ruler per unit, sized from this data - not inherited from a layout
    # measured against figures a hundred times smaller.
    px = t.tree.scale_px
    groups = {n.key: n.scale_group for n in t.tree.nodes}
    heights = {k: (max(v[k]) - min(v[k])) * px[groups[k]] * 0.75 for k in v}
    check("every box is a sane height at its group's scale",
          all(20.0 <= h <= 400.0 for h in heights.values()),
          ", ".join(f"{k} {h:.0f}pt" for k, h in heights.items()))

    check("the tree declares its own rulers",
          set(px) == set(groups.values()),
          f"{sorted(px)}")

    stated = {n.scenario for n in t.provenance}
    check("every scenario drawn declares its provenance",
          set(t.category_scenarios) <= stated, f"stated={sorted(stated)}")


def _c08h(t: D.Template, check) -> None:
    change = t.tier("change").merged()
    level = t.tier("level").merged()
    flow = t.tier("flow").merged()
    outflow = t.tier("outflow").merged()
    n = len(t.categories)
    filed = A.C08H_VALUES["filed_levels"]
    actuals = len(A.C08H_ACTUALS)

    check("every series is the full category length",
          all(len(x) == n for x in (change, level, flow, outflow)),
          f"{n} categories")

    check("no period is drawn twice or left empty",
          all(sum(v is not None for v in
                  (s.values[i] for s in t.tier("level").series)) == 1
              for i in range(n)),
          "one scenario per quarter across all four tiers")

    worst = max(abs((i - o) - c) for i, o, c in zip(flow, outflow, change))
    check("change is losses incurred less claims paid",
          worst < 1e-9, f"largest disagreement {worst:.2e}")

    running = t.opening
    worst = 0.0
    for delta, drawn in zip(change, level):
        running += delta
        worst = max(worst, abs(running - drawn))
    check("level is the opening balance plus every change since",
          worst < 1e-9, f"largest disagreement {worst:.2e} over {n} quarters")

    # The reconciliation, and the reason this template is worth building on
    # real figures: the level is *derived* from two flows, and it has to land
    # on the balance Progressive independently reports at each quarter end. A
    # self-consistent roll-forward proves only that the arithmetic ran.
    worst = max(abs(a - b) for a, b in zip(level[:actuals], filed))
    check("every filed quarter's derived level equals the reported balance",
          worst <= 0.5,
          f"{actuals} quarters, largest difference {worst:,.2f} $m")

    check("the opening balance is the reported one, not a plug",
          abs((filed[0] - change[0]) - t.opening) < 1e-9,
          f"{t.opening:,.0f} $m at the quarter before the first drawn")

    # The plan basis, asserted rather than described. Every constructed
    # quarter must be the stated rule applied to the stated rate - if the
    # sheet says losses run at the target loss ratio, they have to.
    v = A.C08H_VALUES
    check("the plan loss ratio is the target less the trailing expense ratio",
          abs((A.TARGET_COMBINED_RATIO - v["expense_ratio"])
              - v["plan_loss_ratio"]) < 1e-9,
          f"{A.TARGET_COMBINED_RATIO:.0f} - {v['expense_ratio']:.2f} = "
          f"{v['plan_loss_ratio']:.2f}")

    worst = max(abs(o - D.excel_round(i * v["pay_rate"]))
                for i, o in zip(flow[actuals:], outflow[actuals:]))
    check("every constructed quarter pays the stated share of what it books",
          worst < 1e-9,
          f"{n - actuals} quarters at {v['pay_rate'] * 100:.1f}%")

    check("the constructed flows grow at the stated rate, none held flat",
          len({round(x, 3) for x in flow[actuals:]}) == n - actuals,
          f"{n - actuals} distinct values, "
          f"{v['growth'] ** 4 * 100 - 100:+.1f}% a year")

    # A stock chart whose level dips below the axis floor draws off the bottom
    # of its own frame, and the bounds are fractions of a span the sheet
    # computes - so check the property, not the constant.
    span = max(max(level), max(flow), abs(min(outflow)), max(outflow))
    bounds = t.tier_bounds.get("level")
    if bounds is None:
        # Said rather than raised. A KeyError here is true and useless; what
        # the reader needs is which constant is wrong and by how much.
        check("the level chart declares axis bounds for this data", False,
              f"no Template.tier_bounds['level'], so the axis falls back to "
              f"LineLayout.maximum - {L.C08H_LINE.maximum:,.0f}, declared in "
              f"data units above an inventory of 22 tons, against a reserve "
              f"reaching {max(level):,.0f} $m. Every series would clamp to "
              f"the frame and the chart would draw flat")
    else:
        lo, hi = bounds
        check("every drawn value sits inside the declared axis",
              lo * span <= -max(outflow) and max(level) <= hi * span,
              f"axis {lo * span:,.0f}..{hi * span:,.0f} $m holds "
              f"level to {max(level):,.0f} and payments to "
              f"{max(outflow):,.0f}")

    check("the message counts the quarters it claims",
          f"in {sum(1 for c in change[:actuals] if c > 0)} of them"
          in t.title.message,
          "recounted from the changes rather than asserted")

    stated = {note.scenario for note in t.provenance}
    check("every scenario drawn declares its provenance",
          set(t.category_scenarios) <= stated, f"stated={sorted(stated)}")

    check("only the constructed scenarios are marked assumed",
          {note.scenario for note in t.provenance
           if note.basis == "assumed"} == {"FC", "PL"},
          "AC is filed; FC and PL are this workbook's construction")


def _c07c(t: D.Template, check) -> None:
    monthly = t.tier("monthly")
    plan = monthly.series_for("PL").values
    month = [a if a is not None else f for a, f in
             zip(monthly.series_for("AC").values,
                 monthly.series_for("FC").values)]
    cum_pl = t.tier("cum_pl").series_for("PL").values
    cum = t.tier("cum_ac").merged()
    cum_fc = t.tier("cum_fc").series_for("FC").values
    cum = [a if a is not None else f for a, f in zip(cum, cum_fc)]
    mat = t.tier("mat").merged()
    n = len(t.categories)
    v = A.C07C_VALUES
    last = A.C07C_LAST_ACTUAL

    check("every series is the full category length",
          all(len(x) == n for x in (plan, month, cum_pl, cum, mat)),
          f"{n} categories")

    check("no month is drawn twice or left empty",
          all(sum(s.values[i] is not None for s in monthly.series
                  if s.scenario != "PL") == 1 for i in range(n)),
          f"{last} actual then {n - last} forecast, no overlap")

    for name, series, source in (("actual and forecast", cum, month),
                                 ("plan", cum_pl, plan)):
        worst, total = 0.0, 0.0
        for value, drawn in zip(source, series):
            total += value
            worst = max(worst, abs(total - drawn))
        check(f"the {name} cumulative is the running sum of its own months",
              worst < 1e-9, f"largest disagreement {worst:.2e}")

    # December is the one month where a year-to-date total and a twelve-month
    # rolling total mean the same thing, so they have to agree exactly. If the
    # moving total were accumulated from anything but these same months, this
    # is where it would show.
    check("the moving annual total meets the cumulative at December",
          abs(mat[-1] - cum[-1]) < 1e-9,
          f"both {mat[-1]:,.0f} $m")

    # Recomputed from the harvest rather than trusted: each total is this month
    # and the eleven before it, which for January reaches back into the prior
    # year - the series the reference had to transcribe because its page did
    # not carry them.
    worst = 0.0
    for i in range(n):
        total = 0.0
        for back in range(12):
            m, y = i + 1 - back, A.C07C_YEAR
            while m < 1:
                m, y = m + 12, y - 1
            total += (month[m - 1] if y == A.C07C_YEAR
                      else A._c07c_result((y, m)))
        worst = max(worst, abs(total - mat[i]))
    check("every moving annual total is its own trailing twelve months",
          worst < 0.51,
          f"largest disagreement {worst:.2f} $m, recomputed from the harvest")

    check("the plan is earned premium at the target margin",
          all(abs(D.excel_round(e * A.TARGET_MARGIN / 100.0) - p) < 1e-9
              for e, p in zip(v["earned"], plan)),
          f"{A.TARGET_MARGIN:.1f}% of earned premium, all {n} months")

    check("the actual months carry the margin the releases report",
          all(abs(A._c07c_result((A.C07C_YEAR, i + 1)) - month[i]) < 0.51
              for i in range(last)),
          f"{last} months of earned premium times (100 - combined ratio)")

    # The callouts are derived from the series, so the sentence is what has to
    # follow the data rather than the other way round.
    for annotation in t.annotations:
        index = annotation.target[1]
        gap = cum[index] - cum_pl[index]
        check(f"the {t.categories[index]} callout states the gap it points at",
              annotation.text == f"{gap:+,.0f}",
              f"{annotation.text} at {t.categories[index]}")

    lo, hi = t.tier_bounds["line"]
    span = max(max(cum), max(mat), max(cum_pl), max(month))
    check("every drawn value sits inside the declared axis",
          lo * span <= min(min(cum), min(month)) and max(mat) <= hi * span,
          f"axis {lo * span:,.0f}..{hi * span:,.0f} $m over a tallest series "
          f"of {span:,.0f}")

    stated = {note.scenario for note in t.provenance}
    check("every scenario drawn declares its provenance",
          set(t.category_scenarios) | {"PL"} <= stated,
          f"stated={sorted(stated)}")

    check("only the constructed scenarios are marked assumed",
          {note.scenario for note in t.provenance
           if note.basis == "assumed"} == {"PL", "FC"},
          "the actual months are derived from two reported figures, not "
          "assumed")


def _walk_checks(t: D.Template, check, span: float, opening: float) -> None:
    """Shared by both bridges: the steps must walk from opening to closing."""
    steps = t.tier("wf").merged()
    total = opening + sum(steps)
    closing = [s for s in t.summary_rows if not s.before and s.stack]
    stated = sum(value for _scenario, value in closing[0].stack)
    check("the steps walk from the opening total to the closing one",
          abs(total - stated) < 0.51,
          f"{opening:,.0f} + {sum(steps):+,.0f} = {total:,.0f}, "
          f"closing bar says {stated:,.0f}")

    lo, hi = t.tier_bounds["wf"]
    # Every level the walk passes through, not just its ends - a step that
    # overshoots and comes back would clear an endpoints-only check.
    level, levels = opening, [opening]
    for step in steps:
        level += step
        levels.append(level)
    check("the bridge stays inside its own window",
          lo * span - 1e-6 <= min(levels) and max(levels) <= hi * span + 1e-6,
          f"window {lo * span:,.0f}..{hi * span:,.0f} holds a walk through "
          f"{min(levels):,.0f}..{max(levels):,.0f}")
    check("the window is not a negative axis under a positive measure",
          lo >= 0.0 or min(levels) < 0.0,
          f"lowest level {min(levels):,.0f}, window floor {lo * span:,.0f}")

    # The invariant that makes the sheet mean anything: the bridge and the
    # columns it bridges must be on one ruler. Their axes are different
    # windows, so equal ranges would be wrong - what has to match is the
    # dollars each point of chart height is worth. Checked here as well as in
    # test_rescale because it needs no Excel, and because choosing a sensible
    # window for each tier independently is exactly how it gets broken.
    name = f"{t.id}{t.variant}"
    extent = {spec.key: spec.plot_extent for spec in L.LAYOUTS[name].tiers}
    per_point = {}
    for key in L.LAYOUTS[name].same_unit_tiers:
        low, high = t.tier_bounds[key]
        per_point[key] = (high - low) * span / extent[key]
    worst = max(per_point.values()) - min(per_point.values())
    check("the bridge and its columns are on one ruler",
          worst < 1e-6,
          ", ".join(f"{k} {v:,.3f} $m/pt" for k, v in per_point.items()))


def _c06f(t: D.Template, check) -> None:
    ac = t.tier("measure").series_for("AC").values
    py = t.tier("measure").series_for("PY").values
    wf = t.tier("wf").merged()
    rel = t.tier("var_rel").merged()
    n = len(t.categories)

    check("every series is the full category length",
          all(len(x) == n for x in (ac, py, wf, rel)), f"{n} categories")

    worst = max(abs((a - p) - v) for a, p, v in zip(ac, py, wf))
    check("each step is this year less last year",
          worst < 1e-9, f"largest disagreement {worst:.2e}")

    worst = max(abs(v / p * 100.0 - r) for v, p, r in zip(wf, py, rel))
    check("each percentage is its own step over its own base",
          worst < 1e-9, f"largest disagreement {worst:.2e}")

    # Variant F *is* the ordering, so it is asserted rather than assumed - and
    # the catch-all is exempt because it is pinned last on purpose.
    named = wf[:-1]
    check("rows are sorted by step size, largest first",
          all(a >= b for a, b in zip(named, named[1:])),
          f"{len(named)} named states, {named[0]:+,.0f} down to "
          f"{named[-1]:+,.0f}")
    check("the catch-all is pinned last however large it is",
          t.categories[-1] == "All other" and wf[-1] > max(named),
          f"All other adds {wf[-1]:+,.0f}, more than any named state")

    opening = sum(value for _s, value in
                  [s.stack[0] for s in t.summary_rows if s.before][0:1])
    _walk_checks(t, check, A._C06F_SPAN, opening)

    # The message makes an arithmetic claim; it has to be true.
    check("the message's claim about the catch-all holds",
          wf[-1] > sum(named),
          f"outside the top ten {wf[-1]:+,.0f} vs all ten named "
          f"{sum(named):+,.0f}")

    check("both years tie to the filing's own total",
          abs(sum(ac) - A._c06f_total_ac) < 1e-9
          and abs(sum(py) - A._c06f_total_py) < 1e-9,
          f"{A._c06f_total_ac:,.0f} and {A._c06f_total_py:,.0f} $m")

    check("no plan is claimed where none is published",
          not any(s.reference == "PL" for s in t.summary_rows)
          and not any(note.scenario == "PL" for note in t.provenance),
          "no PL row, no PL provenance note")

    stated = {note.scenario for note in t.provenance}
    check("every scenario drawn declares its provenance",
          set(t.category_scenarios) | {"PY"} <= stated,
          f"stated={sorted(stated)}")


def _c05x(t: D.Template, check) -> None:
    measure = t.tier("measure")
    plan = measure.series_for("PL").values
    drawn = [a if a is not None else f for a, f in
             zip(measure.series_for("AC").values,
                 measure.series_for("FC").values)]
    wf = t.tier("wf").merged()
    rel = t.tier("var_rel").merged()
    n = len(t.categories)

    check("every series is the full category length",
          all(len(x) == n for x in (plan, drawn, wf, rel)), f"{n} categories")

    # Only the scenarios that partition the timeline. The reference series -
    # prior year and plan - run the full year alongside them by design.
    check("no month is drawn twice or left empty",
          all(sum(s.values[i] is not None for s in measure.series
                  if s.scenario in ("AC", "FC")) == 1 for i in range(n)),
          f"{A.C05X_LAST_ACTUAL} actual then {n - A.C05X_LAST_ACTUAL} forecast")

    # From the data module: the sheet derives the per-month reference as
    # measure less variance rather than drawing it as a fourth bar.
    prior = list(A.C05X_PY_VALUES)
    worst = max(abs((m - p) - v) for m, p, v in zip(drawn, prior, wf))
    check("each step is the month less the same month last year",
          worst < 1e-9, f"largest disagreement {worst:.2e}")

    worst = max(abs(v / p * 100.0 - r) for v, p, r in zip(wf, prior, rel))
    check("each percentage is its own step over its own base",
          worst < 1e-9, f"largest disagreement {worst:.2e}")

    # The bridge walks from whichever opening column carries its reference, and
    # it reads that column from the row directly above the months - so the
    # seeding column has to be the last of the pair.
    openings = [row for row in t.summary_rows if row.before]
    reference = t.tier("wf").reference
    check("the prior year is drawn once, on the opening column",
          not any(sr.scenario == "PY" for sr in measure.series),
          "three bars a month, not four")
    check("the opening column the bridge walks from sits next to the months",
          openings[-1].stack[0][0] == reference,
          f"walks from {reference}, last opening column is "
          f"{openings[-1].label!r}")

    check("the plan is earned premium at the target margin",
          all(abs(D.excel_round(e * A.TARGET_MARGIN / 100.0) - p) < 1e-9
              for e, p in zip(A.C07C_VALUES["earned"], plan)),
          f"{A.TARGET_MARGIN:.1f}% of earned premium, all {n} months")

    # The point of building this sheet from C07C's series rather than its own:
    # two notations over one set of figures cannot drift apart.
    check("the months are C07C's months, not a second calculation of them",
          tuple(drawn) == tuple(A.C07C_VALUES["monthly"])
          and tuple(plan) == tuple(A.C07C_VALUES["plan"]),
          "identical to the cumulative sheet's own series")

    _walk_checks(t, check, A._C05X_SPAN, A.C05X_PY_TOTAL)

    check("the closing bar splits actual from forecast",
          [s for s in t.summary_rows if not s.before and s.stack][0].stack
          == (("AC", A._c05x_ac), ("FC", A._c05x_fc)),
          f"{A._c05x_ac:,.0f} measured, {A._c05x_fc:,.0f} expected")

    check("the message names the month that actually moves most",
          f"{D.MONTHS[max(range(n), key=lambda i: abs(wf[i]))]} contributing"
          in t.title.message,
          "recounted from the steps rather than asserted")

    stated = {note.scenario for note in t.provenance}
    check("every scenario drawn declares its provenance",
          set(t.category_scenarios) | {"PL", "PY"} <= stated,
          f"stated={sorted(stated)}")

    check("only the constructed scenarios are marked assumed",
          {note.scenario for note in t.provenance
           if note.basis == "assumed"} == {"PL", "FC"},
          "the reported months are derived from two filed figures")


def _t01b(t: D.Template, check) -> None:
    n = len(t.categories)
    elements = [i for i, row in enumerate(t.rows) if row.kind == "element"]

    check("every tier is the full row length",
          all(len(tier.merged()) == n for tier in t.tiers),
          f"{n} rows across {len(t.tiers)} tiers")

    check("the table carries both periods and both references",
          {tier.key for tier in t.tiers} == {
              f"{stem}_{suffix}" for suffix in ("month", "ytd")
              for stem in ("m", "dpy", "dpyp", "dpl", "dplp")},
          "month and year to date, each against prior year and plan")

    # A table has to add up. The subtotals are derived from the elements
    # precisely so that they do, on every scenario and in both periods.
    for suffix, block in (("month", A.T01B_MONTH_BLOCK),
                          ("ytd", A.T01B_YTD_BLOCK)):
        for scenario in ("AC", "PY", "PL"):
            values = block[scenario]
            personal = sum(values[i] for i in (0, 1, 2))
            total = sum(values[i] for i in elements)
            check(f"{suffix} {scenario}: the hierarchy closes",
                  abs(personal - values[3]) < 1e-9
                  and abs(total - values[6]) < 1e-9,
                  f"parts {personal:,.0f} = Personal Lines {values[3]:,.0f}; "
                  f"elements {total:,.0f} = Companywide {values[6]:,.0f}")

    for tier in t.tiers:
        if tier.kind != "variance_abs":
            continue
        suffix = tier.key.split("_", 1)[1]
        block = A.T01B_MONTH_BLOCK if suffix == "month" else A.T01B_YTD_BLOCK
        worst = max(abs((a - r) - v) for a, r, v in
                    zip(block["AC"], block[tier.reference], tier.merged()))
        check(f"{tier.key} is actual less {tier.reference}",
              worst < 1e-9, f"largest disagreement {worst:.2e}")

    for tier in t.tiers:
        if tier.kind != "variance_rel":
            continue
        suffix = tier.key.split("_", 1)[1]
        block = A.T01B_MONTH_BLOCK if suffix == "month" else A.T01B_YTD_BLOCK
        reference = block[tier.reference]
        drawn = tier.merged()
        worst = max((abs((a - r) / r * 100.0 - d)
                     for a, r, d in zip(block["AC"], reference, drawn)
                     if r > 0 and d is not None), default=0.0)
        check(f"{tier.key} is its own variance over its own base",
              worst < 1e-9, f"largest disagreement {worst:.2e}")
        check(f"{tier.key} draws nothing on a non-positive base",
              all(d is None for r, d in zip(reference, drawn) if r <= 0),
              f"{sum(1 for r in reference if r <= 0)} non-positive base(s)")

    # The message ranks the columned segments; the residual is not one.
    best = max(A._T01B_COLUMNED,
               key=lambda i: (A.T01B_YTD_BLOCK["AC"][i]
                              - A.T01B_YTD_BLOCK["PL"][i]))
    check("the message names the segment that really leads",
          t.categories[best] in t.title.message and "Other" != t.categories[best],
          f"{t.categories[best]}, chosen from the columned segments only")
    check("the message's claim that every columned segment beat target holds",
          all(A.T01B_YTD_BLOCK["AC"][i] > A.T01B_YTD_BLOCK["PL"][i]
              for i in A._T01B_COLUMNED),
          f"{len(A._T01B_COLUMNED)} columned segments, all above target")

    stated = {note.scenario for note in t.provenance}
    check("every scenario drawn declares its provenance",
          {"AC", "PY", "PL"} <= stated, f"stated={sorted(stated)}")


def _statement(t: D.Template, check) -> None:
    """Shared by T03A and T04A: they draw one statement, so they share checks."""
    n = len(t.categories)
    values = A.T03A_VALUES

    check("every tier is the full row length",
          all(len(tier.merged()) == n for tier in t.tiers),
          f"{n} rows across {len(t.tiers)} tiers")

    for scenario in ("AC", "PY", "PL"):
        v = values[scenario]
        premium, losses, acquisition, other, expenses, result = v[:6]
        check(f"{scenario}: the expense lines sum to the expense subtotal",
              abs((losses + acquisition + other) - expenses) < 0.51,
              f"{losses:,.0f} + {acquisition:,.0f} + {other:,.0f} = "
              f"{expenses:,.0f}")
        check(f"{scenario}: premium less every expense is the result",
              abs((premium - expenses) - result) < 0.51,
              f"{premium:,.0f} - {expenses:,.0f} = {result:,.0f}")
        check(f"{scenario}: each ratio is its own line over premium",
              all(abs(v[i] - part / premium * 100.0) < 0.06 for i, part in
                  ((6, losses), (7, expenses))),
              f"loss {v[6]:.1f}, combined {v[7]:.1f}")
        check(f"{scenario}: the combined ratio less losses is the expense one",
              abs((v[7] - v[6]) - (acquisition + other) / premium * 100.0)
              < 0.11,
              f"{v[7]:.1f} - {v[6]:.1f} = {v[7] - v[6]:.1f}% of premium")

    # The plan basis, asserted rather than described: the whole point of the
    # plan column is that it lands on the stated target.
    check("the plan lands exactly on the stated combined ratio",
          abs(values["PL"][7] - A.TARGET_COMBINED_RATIO) < 0.05,
          f"{values['PL'][7]:.1f} against a {A.TARGET_COMBINED_RATIO:.0f} "
          f"target")
    check("the plan holds the expense ratio where it actually landed",
          abs(A._t03a_expense("PL") - A._t03a_expense("AC")) < 0.05,
          f"{A._t03a_expense('PL'):.1f}% in both columns, so the target falls "
          f"wholly on losses")
    check("only premium is shared between actual and plan",
          values["PL"][0] == values["AC"][0]
          and values["PL"][1] != values["AC"][1],
          "a combined ratio target says nothing about how much to write")

    for tier in t.tiers:
        if tier.kind not in ("variance_abs", "variance_rel"):
            continue
        reference = values[tier.reference]
        drawn = tier.merged()
        if tier.kind == "variance_abs":
            worst = max(abs((a - r) - d) for a, r, d
                        in zip(values["AC"], reference, drawn))
        else:
            worst = max((abs((a - r) / r * 100.0 - d) for a, r, d
                         in zip(values["AC"], reference, drawn)
                         if r > 0 and d is not None), default=0.0)
        check(f"{tier.key} agrees with the columns it is drawn from",
              worst < 1e-9, f"largest disagreement {worst:.2e}")

    check("the message's claim about where the gain came from holds",
          values["AC"][6] < values["PY"][6]
          and A._t03a_expense("AC") > A._t03a_expense("PY"),
          f"loss ratio {values['PY'][6]:.1f} to {values['AC'][6]:.1f}, "
          f"expense ratio {A._t03a_expense('PY'):.1f} to "
          f"{A._t03a_expense('AC'):.1f}")

    stated = {note.scenario for note in t.provenance}
    check("every scenario drawn declares its provenance",
          {"AC", "PY", "PL"} <= stated, f"stated={sorted(stated)}")


def _t03a(t: D.Template, check) -> None:
    _statement(t, check)
    check("the statement carries both references",
          {tier.key for tier in t.tiers} == {"m", "dpy", "dpyp", "dpl", "dplp"},
          "prior year and plan, in figures only")


def _t04a(t: D.Template, check) -> None:
    _statement(t, check)
    check("the bar sheet drops the prior-year pair",
          {tier.key for tier in t.tiers} == {"m", "dpl", "dplp"},
          "no room for four variance columns beside two panels")
    check("T03A and T04A draw the same statement",
          t.categories == A.T03A.categories
          and t.tier("m").series_for("AC").values
          == A.T03A.tier("m").series_for("AC").values,
          "one block, two sheets")

    # The panel ruler is the template's own, sized from these variances - the
    # layout's is pixels per kEUR against a statement a thousand times smaller.
    _bar_fit(t, check, L.T04A_TABLE.panel_value_width)


def _bar_fit(t: D.Template, check, panel_value_width: float) -> None:
    """Every bar inside its panel, and no panel wider than it needs to be.

    Each side of zero is checked on its own. A single "longest bar fits the
    shorter side" test was what let these panels keep the reference's zero
    position - two-thirds of the way across, for adverse variances this data
    barely has - and that blank two-thirds is what made the columns too wide.
    """
    # Geometry without a panel tier is the silent failure: the column exists,
    # holds the right number, and prints it. The sheet looks finished and is
    # not the template.
    check("every panel with geometry is actually drawn as one",
          set(t.panel_geometry) == set(t.panel_tiers),
          f"{len(t.panel_tiers)} tiers drawn as bars")

    for key, panel in sorted(t.panel_geometry.items()):
        drawn = [v for v in t.tier(key).merged() if v is not None]
        left = max(0.0, -min(drawn)) * panel.scale
        right = max(0.0, max(drawn)) * panel.scale
        check(f"the {key} bars stay inside their panel",
              left <= panel.zero_px + 1e-6
              and right <= panel.width_px - panel.zero_px + 1e-6,
              f"{left:.0f}px left of a zero at {panel.zero_px}, "
              f"{right:.0f}px right of it in {panel.width_px}px")
        column = panel_value_width * panel.width_px / 64.0
        check(f"the {key} column is no wider than the bars need",
              column <= 18.05,
              f"{column:.1f} characters wide")


def _t02a(t: D.Template, check) -> None:
    n = len(t.categories)
    elements = [i for i, row in enumerate(t.rows) if row.kind == "element"]

    check("the bar table drops the prior-year pair",
          {tier.key for tier in t.tiers} == {
              "m_month", "dpl_month", "dplp_month",
              "m_ytd", "dpl_ytd", "dplp_ytd"},
          "measure and plan variance, month and year to date")

    check("T01B and T02A draw the same table",
          t.categories == A.T01B.categories
          and t.tier("m_ytd").series_for("AC").values
          == A.T01B.tier("m_ytd").series_for("AC").values,
          "two blocks, two sheets")

    for suffix, block in (("month", A.T01B_MONTH_BLOCK),
                          ("ytd", A.T01B_YTD_BLOCK)):
        for scenario in ("AC", "PL"):
            values = block[scenario]
            check(f"{suffix} {scenario}: the hierarchy closes",
                  abs(sum(values[i] for i in (0, 1, 2)) - values[3]) < 1e-9
                  and abs(sum(values[i] for i in elements) - values[6]) < 1e-9,
                  f"Personal Lines {values[3]:,.0f}, "
                  f"Companywide {values[6]:,.0f}")
        absolute = t.tier(f"dpl_{suffix}").merged()
        worst = max(abs((a - p) - v) for a, p, v
                    in zip(block["AC"], block["PL"], absolute))
        check(f"dpl_{suffix} is actual less plan",
              worst < 1e-9, f"largest disagreement {worst:.2e}")

    # The claim the notation makes: a month's miss and the year's are drawn at
    # the same pixels per dollar, so they can be compared by eye. Two panels of
    # different widths keep that only if one ruler drives both.
    groups: dict[str, set] = {}
    for key, panel in t.panel_geometry.items():
        groups.setdefault(panel.scale_group, set()).add(round(panel.scale, 9))
    for group, scales in groups.items():
        check(f"the {group} panels share one ruler",
              len(scales) == 1,
              f"{len(t.panel_geometry)} panels, "
              f"{sorted(scales)} px per unit")

    _bar_fit(t, check, L.T02A_TABLE.panel_value_width)

    check("the panel ruler is the template's own, not the layout's",
          set(t.panel_geometry) == {"dpl_month", "dplp_month",
                                    "dpl_ytd", "dplp_ytd"},
          "the layout's is declared in pixels per kEUR")

    stated = {note.scenario for note in t.provenance}
    check("every scenario drawn declares its provenance",
          {"AC", "PL"} <= stated, f"stated={sorted(stated)}")


CHECKS = {"C01A": [_c01a], "C02A": [_c02a], "C03A": [_c03a], "C04A": [_c04a],
          "C09C": [_c09c], "C10D": [_c10d], "C11A": [_c11a], "C12A": [_c12a],
          "C13D": [_c13d], "C08H": [_c08h], "C07C": [_c07c],
          "C06F": [_c06f], "C05X": [_c05x], "T01B": [_t01b], "T03A": [_t03a], "T04A": [_t04a], "T02A": [_t02a]}


def check_template(template: D.Template) -> None:
    """Gate one template, raising on failure. Passed to ``ibcs_excel.build``.

    Same shape as ``ibcs_data.check_ties`` so it can stand in for it, and
    deliberately loud about a template it does not know: a dataset that grows a
    sheet without growing its tie-outs would otherwise be built on nothing.
    """
    name = f"{template.id}{template.variant}"
    if name not in CHECKS:
        raise TieError(f"no tie-outs written for {name}; refusing to build it")
    if run(template, CHECKS[name]):
        raise TieError(f"tie-outs failed for {name}")


def main() -> int:
    total = 0
    for name, template in A.TEMPLATES.items():
        print(f"{name}:")
        if name not in CHECKS:
            # The same refusal check_template makes, said plainly rather than
            # as a KeyError: a template without tie-outs is not verified, and
            # the run should say so rather than crash.
            print(f"  no tie-outs written for {name} - refusing to pass it")
            total += 1
            continue
        total += run(template, CHECKS[name])
    if total:
        print(f"\n{total} tie-out(s) FAILED", file=sys.stderr)
        return 1
    print("\nPASS: every derived figure agrees with what it was derived from.")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    raise SystemExit(main())
