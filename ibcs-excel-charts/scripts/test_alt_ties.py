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


CHECKS = {"C01A": [_c01a], "C02A": [_c02a], "C03A": [_c03a],
          "C10D": [_c10d], "C12A": [_c12a]}


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
