"""Reconcile the harvested monthly figures against Progressive's own totals.

This is the check that invented data could never have had, and the reason using
real figures is worth the extraction work. Three independent tests, each of
which a transcription error has to survive:

**Months sum to the year.** Twelve monthly net-premiums-written figures, pulled
from twelve separate news releases, must add up to the annual figure Progressive
reported in its 10-K and tagged in XBRL - a completely different document,
audited, filed months later. One mistyped month breaks it.

**Each release restates the year before.** Every monthly release prints the same
month a year earlier beside the current one. That prior-year column must equal
what the release of twelve months ago reported as its actual. Two independent
statements of the same fact, twelve months apart.

**Segments add to companywide.** Agency plus Direct plus Property is the
Personal Lines total; Personal Lines plus Commercial is companywide - and
companywide must equal the headline block at the top of the same page.

A tolerance is allowed on the annual sum because Progressive rounds monthly
figures to the nearest million and the 10-K figure is rounded independently;
twelve roundings can legitimately differ from one. It is deliberately tight
enough that a wrong month cannot hide in it.
"""
from __future__ import annotations

import csv
import json
import sys
from pathlib import Path

import fetch_pgr as F

DATA = Path(__file__).resolve().parents[1] / "datasets" / "pgr"

# Twelve figures each rounded to a million can differ from one figure rounded
# to a million by at most six million, plus room for the tenth-of-a-million
# reporting the older releases used. A month is worth billions, so nothing real
# fits in here.
ANNUAL_TOLERANCE = 12.0


def load(name: str) -> list[dict]:
    with (DATA / name).open(encoding="utf-8") as fh:
        return list(csv.DictReader(fh))


def num(row: dict, key: str) -> float | None:
    v = row.get(key, "")
    if v in ("", None, "None"):
        return None
    return float(v)


def annual_npw_from_xbrl() -> dict[int, float]:
    """Progressive's own annual net premiums written, as tagged in its 10-K."""
    raw = F.fetch("https://data.sec.gov/api/xbrl/companyfacts/CIK0000080661.json",
                  cache_name="companyfacts.json")
    facts = json.loads(raw)["facts"]["us-gaap"]["PremiumsWrittenNet"]["units"]["USD"]
    out: dict[int, float] = {}
    for f in facts:
        start, end = f.get("start"), f.get("end")
        if not start or not end:
            continue
        if start[5:] != "01-01" or end[5:] != "12-31" or start[:4] != end[:4]:
            continue
        # Later filings restate the same year; the 10-K for that year wins, and
        # where several agree it does not matter which is taken.
        out[int(end[:4])] = f["val"] / 1e6
    return out


def main() -> int:
    head = load("monthly_headline.csv")
    segs = load("monthly_segments.csv")
    problems: list[str] = []
    passes = 0

    by_month = {(int(r["year"]), int(r["month"])): r for r in head}

    # 1. Twelve months must sum to the audited annual figure.
    annual = annual_npw_from_xbrl()
    print("year   months  sum of months      10-K annual     difference")
    for year in sorted({int(r["year"]) for r in head}):
        months = [r for r in head if int(r["year"]) == year]
        if len(months) != 12 or year not in annual:
            continue
        total = sum(num(r, "npw") or 0.0 for r in months)
        want = annual[year]
        diff = total - want
        ok = abs(diff) <= ANNUAL_TOLERANCE
        passes += ok
        print(f"{year}     {len(months):2}   {total:14,.1f}   {want:14,.1f}   "
              f"{diff:+10,.1f}  {'OK' if ok else 'MISMATCH'}")
        if not ok:
            problems.append(f"{year}: months sum to {total:,.1f} but the 10-K "
                            f"reports {want:,.1f} ({diff:+,.1f})")

    # 2. Every release's prior-year column must equal that month's own actual.
    restated = mismatched = 0
    for (year, month), row in sorted(by_month.items()):
        py = num(row, "npw_py")
        prior = by_month.get((year - 1, month))
        if py is None or prior is None:
            continue
        ac = num(prior, "npw")
        if ac is None:
            continue
        restated += 1
        # The older releases report to a tenth of a million and the later ones
        # to a million, so a restatement can round differently across the change.
        if abs(py - ac) > 1.0:
            mismatched += 1
            problems.append(
                f"{year}-{month:02d} restates {year-1}-{month:02d} NPW as "
                f"{py:,.1f}, but that month's own release said {ac:,.1f}")
    print(f"\nprior-year restatements checked: {restated}, "
          f"disagreeing: {mismatched}")
    passes += restated - mismatched

    # 3. Segments must add up, and match the headline on the same page.
    add_checked = add_bad = 0
    seg_by_month: dict[tuple[int, int], dict[str, dict]] = {}
    for r in segs:
        key = (int(r["year"]), int(r["month"]))
        seg_by_month.setdefault(key, {})[r["role"]] = r
    for key, roles in sorted(seg_by_month.items()):
        cw = roles.get("companywide")
        if not cw:
            continue
        personal = roles.get("personal_total")
        if not personal:
            continue
        # Property counts beside Personal Lines only in the years it was
        # reported outside it; afterwards it is already inside the total.
        outside = [r for r in roles.values()
                   if r["role"] in ("commercial", "property")
                   and r.get("inside_personal", "") != "yes"]
        got = (num(personal, "npw") or 0.0) + sum(num(r, "npw") or 0.0
                                                  for r in outside)
        want = num(cw, "npw")
        if want is None:
            continue
        add_checked += 1
        if abs(got - want) > 1.0:
            add_bad += 1
            problems.append(f"{key[0]}-{key[1]:02d}: personal + commercial = "
                            f"{got:,.1f}, companywide says {want:,.1f}")
        headline_row = by_month.get(key)
        if headline_row is not None:
            h = num(headline_row, "npw")
            if h is not None and abs(h - want) > 1.0:
                add_bad += 1
                problems.append(
                    f"{key[0]}-{key[1]:02d}: headline NPW {h:,.1f} disagrees "
                    f"with the segment table's companywide {want:,.1f}")
    print(f"segment additivity checked: {add_checked}, failing: {add_bad}")
    passes += add_checked - add_bad

    # 4. The year-to-date block must behave like one, and December's must be
    #    the whole year. This is what proves the block taken off the page is
    #    the year to date rather than the quarter sitting beside it - the
    #    mistake that silently cost three years of segment margins while every
    #    premium check still passed.
    ytd = load("ytd_segments.csv")
    ytd_by: dict[tuple[int, int], dict] = {}
    for r in ytd:
        key = (int(r["year"]), int(r["through_month"]))
        ytd_by.setdefault(key, {})[r["role"]] = r

    grew = shrank = 0
    for key, roles in sorted(ytd_by.items()):
        cw, month = roles.get("companywide"), by_month.get(key)
        if not cw or not month:
            continue
        y, m = num(cw, "npw"), num(month, "npw")
        if y is None or m is None:
            continue
        grew += 1
        if y < m - 1.0:
            shrank += 1
            problems.append(f"{key[0]}-{key[1]:02d}: year-to-date premium "
                            f"{y:,.1f} is below the month's own {m:,.1f}")
    print(f"year-to-date blocks checked: {grew}, below their own month: {shrank}")
    passes += grew - shrank

    dec_ok = dec_bad = 0
    for year, want in sorted(annual.items()):
        roles = ytd_by.get((year, 12))
        if not roles or "companywide" not in roles:
            continue
        got = num(roles["companywide"], "npw")
        if got is None:
            continue
        if abs(got - want) > ANNUAL_TOLERANCE:
            dec_bad += 1
            problems.append(f"{year}: December year-to-date says {got:,.1f}, "
                            f"the 10-K says {want:,.1f}")
        else:
            dec_ok += 1
    print(f"December year-to-date vs the 10-K: {dec_ok} match, {dec_bad} differ")
    passes += dec_ok

    # 5. Reported, not asserted: the early years genuinely carry no segment
    #    ratios, which is a fact about the archive rather than a failure. The
    #    annual templates need to know how far back their source reaches.
    have = sorted({int(r["year"]) for r in ytd
                   if int(r["through_month"]) == 12
                   and r["role"] in ("agency", "direct", "property", "commercial")
                   and r["combined_ratio"]})
    if have:
        runs, start = [], have[0]
        for a, b in zip(have, have[1:] + [None]):
            if b != a + 1:
                runs.append((start, a))
                start = b
        lo, hi = max(runs, key=lambda r: r[1] - r[0])
        print(f"full-year segment margins available: {lo}..{hi} "
              f"({hi - lo + 1} consecutive years)")

    # 5b. Year-to-date segment additivity - which went unmeasured while the
    #     monthly blocks were checked, and is what C01A and C02A are built on.
    #     Reported rather than failed: the residual is Progressive's, not ours.
    #     Its releases footnote small run-off businesses that no column carries,
    #     so in some years the named segments genuinely fall short of the
    #     companywide total - 4.3 $m short in 2021, at one-decimal precision, so
    #     not rounding. A chart that claims its parts make the whole has to show
    #     that remainder rather than absorb it.
    resid_checked = 0
    biggest = 0.0
    biggest_year = None
    for (year, month), roles in sorted(ytd_by.items()):
        if month != 12:
            continue
        cw, personal = roles.get("companywide"), roles.get("personal_total")
        if not cw or not personal:
            continue
        outside = sum(num(r, "npw") or 0.0 for r in roles.values()
                      if r["role"] in ("commercial", "property", "other")
                      and r.get("inside_personal", "") != "yes")
        total = (num(personal, "npw") or 0.0) + outside
        want = num(cw, "npw")
        if want is None:
            continue
        resid_checked += 1
        if abs(total - want) > biggest:
            biggest, biggest_year = abs(total - want), year
    if resid_checked:
        print(f"full-year segment residual: checked {resid_checked} years, "
              f"largest unexplained remainder {biggest:,.1f} $m "
              f"({biggest_year}) - shown as an Other segment, not absorbed")
        passes += resid_checked

    # 6. The state table, which is the only figure in the set that comes from an
    #    audited filing rather than a news release - so tying it to the releases
    #    is a check across two different kinds of document.
    states: dict[int, dict[str, float]] = {}
    try:
        for r in load("states.csv"):
            states.setdefault(int(r["year"]), {})[r["state"]] = float(r["npw"])
    except FileNotFoundError:
        states = {}

    add_ok = add_bad = tie_ok = tie_bad = 0
    for year, table in sorted(states.items()):
        total = table.get("total")
        if total is None:
            continue
        parts = sum(v for k, v in table.items() if k != "total")
        if abs(parts - total) > 1.0:
            add_bad += 1
            problems.append(f"{year}: the states plus All other come to "
                            f"{parts:,.1f}, the table's own total says "
                            f"{total:,.1f}")
        else:
            add_ok += 1

        # The 10-K's state total against the same year built from twelve
        # monthly news releases. Two documents, different authors, months apart.
        want = annual.get(year)
        if want is None:
            continue
        if abs(total - want) > ANNUAL_TOLERANCE:
            tie_bad += 1
            problems.append(f"{year}: the state table totals {total:,.1f}, the "
                            f"10-K's tagged annual premium says {want:,.1f}")
        else:
            tie_ok += 1
    if states:
        print(f"state tables: {add_ok} add up, {add_bad} do not; "
              f"{tie_ok} tie to the annual figure, {tie_bad} do not")
        passes += add_ok + tie_ok

    print(f"\n{passes} checks passed, {len(problems)} problem(s)")
    if problems:
        print("\nproblems:", file=sys.stderr)
        for p in problems[:40]:
            print(f"  {p}", file=sys.stderr)
        if len(problems) > 40:
            print(f"  ... and {len(problems) - 40} more", file=sys.stderr)
        return 1
    print("PASS: the harvested months reconcile to Progressive's own totals.")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    raise SystemExit(main())
