"""Harvest The Progressive Corporation's reported figures from SEC EDGAR.

The second workbook is drawn from a real company's published numbers rather
than from invented ones, which means every figure in it has to be traceable to
a filing. Hand-typing them would break that the first time somebody asked
"where did this come from" - so the figures arrive here, from the source, and
the dataset module reads what this writes.

Three things about EDGAR that are not obvious and cost an afternoon each:

**A declared User-Agent is mandatory.** SEC's fair-access policy requires every
request to name a real person and a contact address. Requests without one are
refused with HTTP 403 - not throttled, refused - and the refusal looks like a
network fault rather than a policy one.

**Ten requests a second is the ceiling.** Exceeding it earns a temporary IP
block, so every fetch here goes through one throttled door.

**Nothing is fetched twice.** Responses land in a cache directory and are read
from disk on any later run. A harvest that has to be re-run because a parser
changed must not re-pull ninety-odd filings to do it.

Usage::

    python fetch_pgr.py --probe          # what is available, fetch nothing much
    python fetch_pgr.py --harvest        # pull filings into the cache
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

# SEC requires a real name and contact address on every request - there is no
# compliant way to fetch anonymously, and a request without one is refused.
#
# Read from the environment rather than written here, because this repository is
# public: agreeing to send an address to SEC in a request header is not agreeing
# to publish it on GitHub, and a contact address baked into a source file gets
# published the moment anyone commits.
#
#     $env:SEC_USER_AGENT = 'Your Name you@example.com'
SEC_USER_AGENT_ENV = "SEC_USER_AGENT"
USER_AGENT = os.environ.get(SEC_USER_AGENT_ENV, "")

CIK = "0000080661"                       # The Progressive Corporation
CIK_SHORT = CIK.lstrip("0")

SUBMISSIONS = f"https://data.sec.gov/submissions/CIK{CIK}.json"
ARCHIVE = f"https://www.sec.gov/Archives/edgar/data/{CIK_SHORT}"

# SEC's published ceiling is ten requests a second. Sit well under it: this is
# a one-off harvest and there is nothing to gain by crowding the limit.
MIN_INTERVAL = 0.15

_last_request = [0.0]


def cache_dir() -> Path:
    """Where raw responses live. Under build/, which the repo already ignores."""
    return Path(__file__).resolve().parents[1] / "build" / "pgr-cache"


class NotFound(Exception):
    """EDGAR has no such document. Expected, and not a reason to stop."""


def fetch(url: str, *, cache_name: str | None = None, binary: bool = False):
    """One throttled, cached GET. Returns text unless ``binary``.

    Failures are classified rather than lumped together, because they mean
    genuinely different things and only one of them is worth stopping for:

    * **403** - the User-Agent was rejected. Fatal: nothing will work until it
      is fixed, and grinding through two hundred filings to say so is waste.
    * **404** - this particular document is not there. Raised as ``NotFound``
      so a caller can try the next candidate.
    * **429 / 5xx** - throttled, or EDGAR having a moment. Retried with a
      widening pause. A single transient 503 killing a two-hundred-filing
      harvest, as it did on the first run, is a defect in this function rather
      than in the network.
    """
    name = cache_name or url.replace("https://", "").replace("/", "_")
    path = cache_dir() / name
    if path.exists():
        return path.read_bytes() if binary else path.read_text(
            encoding="utf-8", errors="replace")

    if not USER_AGENT:
        raise SystemExit(
            f"set {SEC_USER_AGENT_ENV} before harvesting - SEC's fair-access "
            f"policy requires every request to name a real person and a "
            f"contact address, and refuses those that do not. "
            f"For example: $env:{SEC_USER_AGENT_ENV} = "
            f"'Your Name you@example.com'")

    req = urllib.request.Request(url, headers={
        "User-Agent": USER_AGENT,
        "Accept-Encoding": "gzip, deflate",
        "Host": url.split("/")[2],
    })
    raw = None
    for attempt in range(5):
        wait = MIN_INTERVAL - (time.monotonic() - _last_request[0])
        if wait > 0:
            time.sleep(wait)
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                raw = resp.read()
            break
        except urllib.error.HTTPError as exc:
            if exc.code == 403:
                raise SystemExit(
                    f"EDGAR rejected the User-Agent on {url} (HTTP 403). "
                    f"SEC requires a real name and contact address; the "
                    f"current one is {USER_AGENT!r}.") from exc
            if exc.code == 404:
                raise NotFound(url) from exc
            if exc.code not in (429, 500, 502, 503, 504) or attempt == 4:
                raise SystemExit(f"EDGAR returned HTTP {exc.code} for {url}") from exc
            time.sleep(2 ** attempt)
        except urllib.error.URLError as exc:
            if attempt == 4:
                raise SystemExit(f"could not reach EDGAR for {url}: {exc}") from exc
            time.sleep(2 ** attempt)
        finally:
            _last_request[0] = time.monotonic()

    if raw[:2] == b"\x1f\x8b":                      # gzip, despite the header
        import gzip
        raw = gzip.decompress(raw)

    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(raw)
    return raw if binary else raw.decode("utf-8", errors="replace")


def submissions() -> dict:
    """The filing index, recent page plus every older page EDGAR splits off."""
    data = json.loads(fetch(SUBMISSIONS, cache_name="submissions.json"))
    recent = data["filings"]["recent"]
    rows = _rows(recent)
    for extra in data["filings"].get("files", []):
        name = extra["name"]
        older = json.loads(fetch(f"https://data.sec.gov/submissions/{name}",
                                 cache_name=name))
        rows += _rows(older)
    return {"name": data.get("name", ""), "rows": rows}


def _rows(block: dict) -> list[dict]:
    """EDGAR ships parallel arrays; a list of dicts is easier to think about."""
    keys = ("accessionNumber", "filingDate", "form", "primaryDocument",
            "reportDate", "items")
    n = len(block.get("accessionNumber", []))
    out = []
    for i in range(n):
        out.append({k: (block.get(k) or [None] * n)[i] for k in keys})
    return out


def probe() -> int:
    """What is actually available, without pulling any filing bodies."""
    data = submissions()
    rows = data["rows"]
    print(f"{data['name']} - CIK {CIK}")
    print(f"{len(rows):,} filings indexed\n")

    by_form: dict[str, list[dict]] = {}
    for r in rows:
        by_form.setdefault(r["form"] or "?", []).append(r)

    print("form      count  earliest     latest")
    for form, group in sorted(by_form.items(),
                              key=lambda kv: -len(kv[1]))[:12]:
        dates = sorted(g["filingDate"] for g in group if g["filingDate"])
        print(f"{form:8} {len(group):6}  {dates[0]}   {dates[-1]}")

    eight_k = by_form.get("8-K", [])
    ten_k = by_form.get("10-K", [])
    print(f"\n8-K filings: {len(eight_k)}")
    print(f"10-K filings: {len(ten_k)}")
    if ten_k:
        print("  10-K report dates: "
              + ", ".join(sorted({r["reportDate"] or "?" for r in ten_k})))
    return 0


# Files in a filing directory that are never the release: EDGAR's own
# generated report fragments, the XBRL sidecars, and the index pages.
NOT_A_RELEASE = re.compile(
    r"(?i)^(r\d+\.htm|report\.css|show\.js|filingsummary\.xml|metalinks\.json)$"
    r"|-index|_(cal|def|lab|pre)\.xml$|\.(xsd|xml|jpg|png|gif|zip|js|css)$")


def release_docs(accession: str) -> list[str]:
    """Candidate URLs for the news release inside one filing, best first.

    Filenames are not stable across fifteen years and the release is not always
    an exhibit. Three eras, all of which have to work:

    * ``pgr20251031ex99earningsrel.htm`` - a numbered EX-99 exhibit (current)
    * ``a8-kaugust2016earningsrele.htm`` - the *primary document* is the
      release, with no "ex99" anywhere in the name (mid-2010s)
    * ``l97809aexv99wa.txt`` - an EX-99 under the old naming (2000s)

    Matching on "ex99" alone silently found nothing for the whole middle era
    and reported it as "Progressive filed no release", which is the sort of
    wrong answer that looks like data. So candidates are ranked rather than
    chosen: anything whose name suggests a release first, then everything else
    by size, and the caller parses down the list until one yields figures.
    """
    a = accession.replace("-", "")
    try:
        idx = json.loads(fetch(f"{ARCHIVE}/{a}/index.json",
                               cache_name=f"idx_{a}.json"))
    except NotFound:
        return []

    scored: list[tuple[int, int, str]] = []
    for item in idx["directory"]["item"]:
        name = item["name"]
        low = name.lower()
        if NOT_A_RELEASE.search(low) or not low.endswith((".htm", ".html", ".txt")):
            continue
        size = int(item.get("size") or 0)
        if size < 4000:                     # a cover page, not a release
            continue
        hint = 0
        if re.search(r"ex-?v?99", low):
            hint -= 2
        if "earning" in low or "release" in low or "result" in low:
            hint -= 2
        scored.append((hint, -size, name))

    scored.sort()
    return [f"{ARCHIVE}/{a}/{n}" for _h, _s, n in scored[:4]]


def annual_report_docs(accession: str) -> list[str]:
    """Candidate URLs for a 10-K's annual-report exhibit, largest first.

    The state table lives in the exhibit rather than in the 10-K body, and the
    exhibit's name is not stable across the archive -
    ``pgr-20251231_d2.htm``, ``exhibit13annualreport2016.htm``,
    ``pgr-20191231exhibit13a.htm``. It is always the substantial document in the
    filing, so size does the finding and the parser confirms it.
    """
    a = accession.replace("-", "")
    try:
        idx = json.loads(fetch(f"{ARCHIVE}/{a}/index.json",
                               cache_name=f"idx_{a}.json"))
    except NotFound:
        return []
    docs = []
    for item in idx["directory"]["item"]:
        name = item["name"]
        if NOT_A_RELEASE.search(name.lower()):
            continue
        if not name.lower().endswith((".htm", ".html")):
            continue
        docs.append((-int(item.get("size") or 0), name))
    docs.sort()
    return [f"{ARCHIVE}/{a}/{n}" for _size, n in docs[:3]]


def harvest_states() -> dict[int, dict[str, float]]:
    """Net premiums written by state, every year the archive reaches.

    Each 10-K carries five years, so the filings overlap heavily. Later filings
    win where they disagree: a restatement is the company correcting itself, and
    the most recent statement of a figure is the one it stands behind.
    """
    import pgr_parse as P

    tens = [r for r in submissions()["rows"]
            if r["form"] == "10-K" and r["reportDate"]
            and r["reportDate"] >= "2010-12-31"]
    tens.sort(key=lambda r: r["reportDate"])

    found: dict[int, dict[str, float]] = {}
    for r in tens:
        a = r["accessionNumber"].replace("-", "")
        for url in annual_report_docs(r["accessionNumber"]):
            try:
                raw = fetch(url, cache_name=f"tenk_{a}_{url.rsplit('/', 1)[-1]}")
            except NotFound:
                continue
            table = P.state_table(P.to_rows(raw))
            if table:
                found.update(table)
                break
    return found


def _plausible(parsed: dict, filed: str) -> bool:
    """Is the parsed period consistent with when the release was filed?

    Progressive reports a month within weeks of its end, so a release filed in
    December 2018 reports on some month between roughly August and December
    2018. Anything else means the period was misread - which is exactly how a
    November 2018 release once got filed under November 2019, with every other
    check still passing because the figures themselves were fine.
    """
    fy, fm = int(filed[:4]), int(filed[5:7])
    months_ago = (fy * 12 + fm) - (parsed["year"] * 12 + parsed["month"])
    return 0 <= months_ago <= 4


def harvest(since: str, until: str) -> int:
    """Pull every monthly release in a date range and report what came back."""
    import pgr_parse as P

    rows = [r for r in submissions()["rows"]
            if r["form"] == "8-K" and r["filingDate"]
            and since <= r["filingDate"] <= until]
    rows.sort(key=lambda r: r["filingDate"])
    print(f"{len(rows)} 8-K filings between {since} and {until}\n")

    found: dict[tuple[int, int], dict] = {}
    skipped = 0
    for r in rows:
        a = r["accessionNumber"].replace("-", "")
        parsed = None
        for url in release_docs(r["accessionNumber"]):
            try:
                raw = fetch(url, cache_name=f"doc_{a}_{url.rsplit('/', 1)[-1]}")
            except NotFound:
                continue
            candidate = P.parse(raw, url.rsplit("/", 1)[-1], r["filingDate"])
            if candidate["year"] is not None and candidate["headline"].get("npw"):
                parsed = candidate
                parsed["source"] = url
                break
        if parsed is None:
            skipped += 1
            continue
        key = (parsed["year"], parsed["month"])
        parsed["accession"] = r["accessionNumber"]
        parsed["filed"] = r["filingDate"]
        # A month reported twice keeps the first release of it. Later filings
        # that mention the same month are quarterly or annual wrap-ups, and
        # their figures are the same ones restated.
        found.setdefault(key, parsed)

    months = sorted(found)
    if not months:
        print(f"parsed nothing from {len(rows)} filings - every one was skipped. "
              f"That is a parser failure, not an empty archive.")
        return 1
    print(f"parsed {len(found)} distinct months "
          f"({months[0][0]}-{months[0][1]:02d} .. "
          f"{months[-1][0]}-{months[-1][1]:02d}), skipped {skipped} filings\n")

    # Coverage by year, because a gap matters more than a count.
    by_year: dict[int, list[int]] = {}
    for y, m in months:
        by_year.setdefault(y, []).append(m)
    print("year  months  missing")
    for y in sorted(by_year):
        have = set(by_year[y])
        missing = [m for m in range(1, 13) if m not in have]
        print(f"{y}    {len(have):2}      "
              f"{', '.join(str(m) for m in missing) if missing else '-'}")

    write_csv(found)
    return 0


def write_csv(found: dict) -> None:
    """One row per month per segment, plus a companywide headline file."""
    out = Path(__file__).resolve().parents[1] / "datasets" / "pgr"
    out.mkdir(parents=True, exist_ok=True)

    head = out / "monthly_headline.csv"
    with head.open("w", encoding="utf-8", newline="") as fh:
        fh.write("year,month,npw,npe,net_income,combined_ratio,"
                 "npw_py,npe_py,net_income_py,combined_ratio_py,"
                 "accession,filed\n")
        for (y, m), p in sorted(found.items()):
            h = p["headline"]
            def g(k, side):
                # A missing figure is an empty cell, never the word "None":
                # the newer releases dropped the prior-year column entirely,
                # and str(None) in a numeric column is a parse error waiting
                # to happen in whatever reads this next.
                if k not in h or h[k][side] is None:
                    return ""
                return h[k][side]
            fh.write(f"{y},{m},{g('npw','ac')},{g('npe','ac')},"
                     f"{g('net_income','ac')},{g('combined_ratio','ac')},"
                     f"{g('npw','py')},{g('npe','py')},"
                     f"{g('net_income','py')},{g('combined_ratio','py')},"
                     f"{p['accession']},{p['filed']}\n")

    seg = out / "monthly_segments.csv"
    with seg.open("w", encoding="utf-8", newline="") as fh:
        # inside_personal travels with the row because it decides what
        # companywide is the sum of, and a reader holding only the
        # figures cannot tell which reporting era a month belongs to.
        fh.write("year,month,role,name,inside_personal,npw,npe,"
                 "loss_ratio,expense_ratio,combined_ratio\n")
        for (y, m), p in sorted(found.items()):
            for c in p["segments"]:
                fh.write(f"{y},{m},{c.get('role','')},{c['name']},"
                         f"{'yes' if c.get('inside_personal') else ''},"
                         f"{c.get('npw','')},{c.get('npe','')},"
                         f"{c.get('loss_ratio','')},"
                         f"{c.get('expense_ratio','')},"
                         f"{c.get('combined_ratio','')}\n")
    # The year-to-date block. December's row is the full financial year, which
    # is the only place annual figures exist *by segment* anywhere in the
    # archive - the 10-K tags premium companywide, not per segment.
    ytd = out / "ytd_segments.csv"
    with ytd.open("w", encoding="utf-8", newline="") as fh:
        fh.write("year,through_month,role,name,inside_personal,npw,npe,"
                 "loss_ratio,expense_ratio,combined_ratio\n")
        for (y, m), rec in sorted(found.items()):
            for c in rec.get("segments_ytd") or []:
                fh.write(f"{y},{m},{c.get('role','')},{c['name']},"
                         f"{'yes' if c.get('inside_personal') else ''},"
                         f"{c.get('npw','')},{c.get('npe','')},"
                         f"{c.get('loss_ratio','')},"
                         f"{c.get('expense_ratio','')},"
                         f"{c.get('combined_ratio','')}\n")

    # Net premiums written by state, from the 10-K annual-report exhibits. The
    # only geographic cut Progressive publishes, and the only figure in the set
    # that comes from an audited filing rather than a news release.
    states = harvest_states()
    if states:
        by_state = out / "states.csv"
        with by_state.open("w", encoding="utf-8", newline="") as fh:
            fh.write("year,state,npw\n")
            for year in sorted(states):
                for name, value in sorted(states[year].items(),
                                          key=lambda kv: -kv[1]):
                    fh.write(f"{year},{name},{value}\n")
        years = sorted(states)
        named = len(states[years[-1]]) - 2      # less "all other" and "total"
        print(f"\nstate table: {len(years)} years ({years[0]}..{years[-1]}), "
              f"{named} states named plus All other")
        print(f"wrote {by_state}")

    print(f"\nwrote {head}")
    print(f"wrote {seg}")
    print(f"wrote {ytd}")


def main(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--probe", action="store_true",
                    help="report what EDGAR holds; fetch only the index")
    ap.add_argument("--harvest", action="store_true",
                    help="pull monthly releases into the cache and write CSVs")
    ap.add_argument("--since", default="2011-01-01")
    ap.add_argument("--until", default="2026-12-31")
    args = ap.parse_args(argv[1:])
    if args.probe:
        return probe()
    if args.harvest:
        return harvest(args.since, args.until)
    ap.print_help()
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main(sys.argv))
