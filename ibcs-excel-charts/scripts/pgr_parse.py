"""Turn one Progressive monthly news release into figures.

The releases are HTML tables, and every one of them carries the same two things
worth having:

**The headline block** - net premiums written and earned, net income, per-share,
combined ratio - for the month, *and for the same month a year earlier*. That
second column is the reason this is worth parsing rather than typing: it makes
every figure checkable against the release that first reported it, twelve months
back. A transcription error has to be made twice, identically, to survive.

**The supplemental block** - the same month split by segment: Agency, Direct,
Property, Personal Lines total, Commercial, Companywide, each with premiums
written, premiums earned and the three GAAP ratios.

That second block is what makes an XY template possible on real data. An
insurer's underwriting profit is earned premium times underwriting margin, and
margin is ``100 - combined ratio`` - so the third measure is the product of the
two axes, which is exactly the relationship C09C is drawn to show. Nothing is
constructed; it is how the industry reports.

Fifteen years of releases are not identically formatted, so parsing is done on
row labels rather than on positions, and every extracted figure is returned with
the label it was found under so a caller can see what it got.
"""
from __future__ import annotations

import html
import re

# Companywide headline rows, as their labels appear in the release.
HEADLINE = {
    "net premiums written": "npw",
    "net premiums earned": "npe",
    "net income": "net_income",
    "combined ratio": "combined_ratio",
}

# The supplemental table's columns, left to right. Property joined Personal
# Lines partway through the period, so the column count is not constant and the
# header row is read rather than assumed.
SEGMENT_ROWS = {
    "net premiums written": "npw",
    "net premiums earned": "npe",
    "loss/lae ratio": "loss_ratio",
    "expense ratio": "expense_ratio",
    "combined ratio": "combined_ratio",
}

MONTHS = ("january february march april may june july august september "
          "october november december").split()


def to_rows(raw: str) -> list[list[str]]:
    """The document as a list of table rows, each a list of cell strings."""
    t = re.sub(r"(?is)<(script|style).*?</\1>", " ", raw)
    t = re.sub(r"(?i)</tr\s*>", "\x00", t)
    t = re.sub(r"(?i)</t[dh]\s*>", "\x01", t)
    t = re.sub(r"(?s)<[^>]+>", "", t)
    t = html.unescape(t)
    rows = []
    for line in t.split("\x00"):
        cells = [re.sub(r"\s+", " ", c).strip() for c in line.split("\x01")]
        cells = [c for c in cells if c not in ("", "$", "%")]
        if cells:
            rows.append(cells)
    return rows


def number(cell: str) -> float | None:
    """One cell as a number, or None where it holds no figure.

    Parentheses are negative - the releases use accounting notation throughout,
    and a loss read as a positive is the kind of error that ties out perfectly
    and means the opposite of the truth.
    """
    c = cell.strip().replace(",", "").replace("$", "").replace("%", "").strip()
    if not c or c in ("-", "—", "–", "NM", "nm"):
        return None
    neg = c.startswith("(") and c.endswith(")")
    if neg:
        c = c[1:-1]
    try:
        v = float(c)
    except ValueError:
        return None
    return -v if neg else v


def period_candidates(rows: list[list[str]],
                      filename: str = "") -> list[tuple[int, int, str]]:
    """Every period this document could be reporting on, best guess first.

    A single answer is not safe to give here. The releases that accompany a
    quarter - filed in January and April - discuss the month, the quarter and
    the year, and mention other periods in their footnotes besides. Committing
    to the first month name found read January 2016's release as *April 2015*.

    So all readings are returned with the evidence they came from, and the
    caller resolves them against the filing date, which is the one fact the
    document cannot be wrong about. Ordered by how much the source is trusted:
    an explicit "month ended" heading, then a fused "December2011" table
    header, then the filename.
    """
    out: list[tuple[int, int, str]] = []

    def add(year: int, month: int, source: str) -> None:
        if 1 <= month <= 12 and (year, month, source) not in out:
            out.append((year, month, source))

    for cells in rows:
        text = " ".join(cells).lower()
        for m in re.finditer(
                r"for the (?:month|period) ended ([a-z]+) \d{1,2},? (\d{4})", text):
            if m.group(1) in MONTHS:
                add(int(m.group(2)), MONTHS.index(m.group(1)) + 1, "month-ended")

    joined = " ".join(" ".join(cells) for cells in rows).lower()
    for m in re.finditer(r"\b(" + "|".join(MONTHS) + r")\s?(20\d{2})\b", joined):
        add(int(m.group(2)), MONTHS.index(m.group(1)) + 1, "table-header")

    low = filename.lower()
    m = re.search(r"(20\d{2})(\d{1,2})(\d{2})", low)
    if m:
        add(int(m.group(1)), int(m.group(2)), "filename")
    m = re.search(r"(" + "|".join(MONTHS) + r")(20\d{2})", low)
    if m:
        add(int(m.group(2)), MONTHS.index(m.group(1)) + 1, "filename")
    return out


def period_of(rows: list[list[str]], filename: str = "") -> tuple[int, int] | None:
    """(year, month) the release reports on.

    Read from the supplemental section's own heading - "For the month ended
    November 30, 2018" - which is the only statement in the document that names
    the period unambiguously. **The whole document is searched**, not the first
    screenful: that heading sits around row 120, below the headline tables, and
    looking only at the top silently found nothing for a third of the archive.

    The filename is the fallback, because it encodes the period end date
    (``pgr20181130exhibit99earnin.htm``, ``a8-knovember2018earningsre.htm``).

    What is deliberately *not* done any more is guessing from the dateline.
    "PROGRESSIVE REPORTS NOVEMBER RESULTS ... December 12, 2018" tempts you to
    take the month from one phrase and the year from another, and doing that
    filed the November 2018 release under November *2019* - a silent
    off-by-one-year that would have put a year's figures in the wrong column
    with every internal check still passing.
    """
    for cells in rows:
        text = " ".join(cells).lower()
        m = re.search(r"for the (?:month|period) ended ([a-z]+) \d{1,2},? (\d{4})",
                      text)
        if m and m.group(1) in MONTHS:
            return int(m.group(2)), MONTHS.index(m.group(1)) + 1

    # Releases before about 2015 carry no "month ended" phrase at all. They do
    # head the policies-in-force table with the month and year run together -
    # "December2011 | December2010" - which is just as unambiguous.
    #
    # The year must follow the month name immediately. That is what separates
    # this from the dateline, "January 19, 2012", which names the month the
    # release was *published* rather than the one it reports: allowing a day
    # number in between would read every December release as a January one.
    joined = " ".join(" ".join(cells) for cells in rows).lower()
    m = re.search(r"\b(" + "|".join(MONTHS) + r")\s?(20\d{2})\b", joined)
    if m:
        return int(m.group(2)), MONTHS.index(m.group(1)) + 1

    low = filename.lower()
    m = re.search(r"(20\d{2})(\d{1,2})(\d{2})", low)      # pgr20181130..., pgr2022430...
    if m:
        year, month = int(m.group(1)), int(m.group(2))
        if 1 <= month <= 12:
            return year, month
    m = re.search(r"(" + "|".join(MONTHS) + r")(20\d{2})", low)
    if m:
        return int(m.group(2)), MONTHS.index(m.group(1)) + 1
    return None


def label_of(cell: str) -> str:
    """A row label, normalised for lookup.

    Footnote markers ride on the label itself - the June 2016 release says
    "Net premiums written1" - and an exact-match lookup misses those rows
    entirely. It then matches the *year-to-date* row of the same name further
    down the page, so the month silently receives a cumulative figure. That is
    how 2016's twelve months came to sum to nine billion more than the 10-K.
    """
    text = cell.strip().lower().rstrip(":")
    text = re.sub(r"\s*\(?\d+\)?$", "", text)      # trailing footnote marker
    return text.strip()


def headline_years(rows: list[list[str]]) -> int:
    """How many year columns the headline table has: two, or only one.

    Not a detail. Until late 2023 the table read
    ``(millions ...) | 2023 | 2022 | Change``; from then on it reads
    ``(millions ...) | 2023`` and the prior-year comparative is gone.

    Code that assumed two columns and demanded two numbers per row skipped the
    whole headline table on the newer releases, then matched the same row
    labels in the *year-to-date* section below - reporting one month's premiums
    as fifty-two billion. So the shape is read rather than assumed.
    """
    for cells in rows[:40]:
        if not cells or "millions" not in cells[0].lower():
            continue
        years = [c.strip() for c in cells[1:] if re.fullmatch(r"20\d{2}", c.strip())]
        # The row repeats for the quarter, so only the leading run counts.
        if len(years) >= 2 and years[0] != years[1]:
            return 2
        if years:
            return 1
    return 2


def headline(rows: list[list[str]], years: int | None = None) -> dict:
    """Companywide figures for the month, and the prior year where given.

    The first match wins. The headline table is at the top of every release and
    the same row labels reappear below it for the quarter and the year, so
    taking anything but the first would take a cumulative figure for a month.
    """
    if years is None:
        years = headline_years(rows)
    out: dict[str, dict] = {}
    for cells in rows:
        key = HEADLINE.get(label_of(cells[0]))
        if not key or key in out:
            continue
        nums = [number(c) for c in cells[1:]]
        nums = [n for n in nums if n is not None]
        if not nums:
            continue
        if years >= 2 and len(nums) >= 2:
            out[key] = {"ac": nums[0], "py": nums[1]}
        elif years == 1:
            out[key] = {"ac": nums[0], "py": None}
    return out


def segment_blocks(rows: list[list[str]]) -> list[list[dict]]:
    """Every supplemental segment block in the release, in page order.

    There are two. The first is the current month; the second repeats the same
    columns for the year to date. Only the first was read for a long time, and
    that threw away the only source of *annual* segment figures in the whole
    archive - which is what a five-year-by-segment template needs, and what
    lets a bubble chart place a business acquired mid-year against a prior year
    it did not exist in.

    Which block is which is taken from page order, because the headings above
    them do not say consistently. That assumption is then **proved** rather than
    trusted: year-to-date premium must exceed the month's, and December's
    year-to-date must equal the annual figure in the 10-K.
    """
    heads = [i for i, cells in enumerate(rows)
             if len(cells) >= 3
             and "agency" in [c.strip().lower() for c in cells]
             and "direct" in [c.strip().lower() for c in cells]]

    blocks: list[list[dict]] = []
    for start in heads:
        cells = rows[start]
        names = [re.sub(r"\d+$", "", c).strip() for c in cells]
        banner = " ".join(" ".join(c)
                          for c in rows[max(0, start - 4):start]).lower()
        cols: list[dict] = [{"name": n, "column": j}
                            for j, n in enumerate(names)]
        for line in rows[start + 1:start + 14]:
            key = SEGMENT_ROWS.get(label_of(line[0]))
            if not key:
                continue
            nums = [n for n in (number(c) for c in line[1:]) if n is not None]
            if len(nums) != len(cols):
                # A row that does not span the block - a catastrophe ratio with
                # blanks in it - cannot be aligned by position, and guessing
                # would file one segment's figure under another's name.
                continue
            for col, value in zip(cols, nums):
                col.setdefault(key, value)
        cols = [c for c in cols if "npw" in c]
        if len(cols) >= 4:
            assign_roles(cols, banner=banner)
            blocks.append(cols)
    return blocks


def segments(rows: list[list[str]]) -> list[dict]:
    """The current-month segment block: the first one on the page.

    A **list** and not a dict. The header row reads
    ``Agency | Direct | Property | Total | Business | Total`` and those two
    Totals are different things - the Personal Lines subtotal and the
    companywide one. Keyed by name, the second silently overwrote the first and
    the companywide column vanished without an error. Position is the only
    thing that distinguishes them, so position is what is kept.
    """
    blocks = segment_blocks(rows)
    return blocks[0] if blocks else []


def assign_roles(cols: list[dict], tol: float = 1.0, banner: str = "") -> None:
    """Name each segment column by what its figures do, not where it sits.

    Position is not enough, because the columns genuinely changed meaning.
    Until recently Property was reported *outside* Personal Lines::

        Agency | Direct | Total(Personal) | Business(Comml) | Business(Prop) | Total

    and it is now reported *inside* it::

        Agency | Direct | Property | Total(Personal) | Business(Comml) | Total

    Two columns are called "Total" and two are called "Business" in both
    layouts, so the labels cannot separate them either. What can is arithmetic:
    Personal Lines is by definition the column equal to the personal leaves
    added up. Reading roles off the sums also makes a mis-assignment impossible
    rather than merely unlikely - where the sums do not work no role is
    assigned at all, and the caller drops the block instead of publishing it.
    """
    def npw(c):
        return c.get("npw")

    def is_property(col: dict) -> bool:
        """Is this leaf the Property business, on the filing's own say-so?

        Two places can say so, because the layouts disagree about which. Where
        Property sits inside Personal Lines its own column is headed
        "Property"; where it sits outside, the column is headed "Business" and
        only the group banner above names it. Either is evidence; neither alone
        is enough.

        Where nothing says Property, the leaf is called "other" rather than
        guessed at. Before the ARX acquisition Progressive reported no Property
        segment and that slot held Other Businesses - figures that add up
        perfectly under a name this code would have invented.
        """
        return "property" in col.get("name", "").lower() or "property" in banner

    if any(npw(c) is None for c in cols):
        return
    cols[0]["role"] = "agency"
    cols[1]["role"] = "direct"
    cols[-1]["role"] = "companywide"

    base = npw(cols[0]) + npw(cols[1])
    mids = cols[2:-1]
    for i, cand in enumerate(mids):
        others = [m for j, m in enumerate(mids) if j != i]
        inside = None
        if abs(npw(cand) - base) > tol:
            inside = next((o for o in others
                           if abs(npw(cand) - (base + npw(o))) <= tol), None)
            if inside is None:
                continue
        cand["role"] = "personal_total"
        if inside is not None:
            inside["role"] = "property" if is_property(inside) else "other"
            # Recorded, because it decides what companywide is the sum of.
            # While Property sat outside Personal Lines, companywide was
            # personal + commercial + property; now it is personal + commercial
            # with property already counted inside the first.
            inside["inside_personal"] = True
        rest = [o for o in others if o is not inside]
        if len(rest) == 1:
            rest[0]["role"] = "commercial"
        elif len(rest) == 2:
            # Commercial and Property, both outside Personal Lines. Commercial
            # has been the larger of the two by premium in every period they
            # were reported apart, and the additivity check below catches it if
            # that ever stops being true.
            big, small = sorted(rest, key=npw, reverse=True)
            big["role"] = "commercial"
            # The smaller leaf is Property only where the filing names it.
            # Before the ARX acquisition Progressive had no Property segment at
            # all, and the column in that slot is "Other Businesses" - real
            # figures under a label this code invented. Arithmetic cannot tell
            # those apart, because both add up perfectly; only the heading can.
            small["role"] = "property" if is_property(small) else "other"
        return


def resolve_period(candidates: list[tuple[int, int, str]],
                   filed: str) -> tuple[int, int, str] | None:
    """Which of the candidate periods the release is actually reporting on.

    Progressive reports a month within weeks of its end, so the answer is the
    most recent candidate that is not *after* the filing date and not more than
    four months before it. Among those, the earliest-listed wins, because the
    list is already ordered by how much its evidence is trusted.
    """
    fy, fm = int(filed[:4]), int(filed[5:7])
    for year, month, source in candidates:
        months_ago = (fy * 12 + fm) - (year * 12 + month)
        if 0 <= months_ago <= 4:
            return year, month, source
    return None


def parse(raw: str, filename: str = "", filed: str = "") -> dict:
    """Everything one release yields, or as much of it as the release carries."""
    rows = to_rows(raw)
    if filed:
        cands = period_candidates(rows, filename)
        resolved = resolve_period(cands, filed)
        period = resolved[:2] if resolved else None
        source = resolved[2] if resolved else ""
    else:
        period = period_of(rows, filename)
        source = "legacy"
    head = headline(rows)
    blocks = segment_blocks(rows)
    segs = blocks[0] if blocks else []

    # The year-to-date block, kept only where it behaves like one. Premium
    # accumulated over a year cannot be smaller than one month of it, so a
    # second block that fails that test is not the year to date and is dropped
    # rather than filed as annual data.
    # The **last** block, not the second. A release covering a quarter end
    # carries three - month, quarter, year to date - and in several years the
    # year to date is split across two tables, premiums in one and the GAAP
    # ratios in the next. Both of those carry the annual premium, so taking
    # the second got the right totals with every ratio blank: the segment
    # margins for 2015, 2019 and 2020 went missing while every premium check
    # still passed.
    ytd = blocks[-1] if len(blocks) > 1 else []
    if ytd and segs:
        m_cw = next((c for c in segs if c.get("role") == "companywide"), None)
        y_cw = next((c for c in ytd if c.get("role") == "companywide"), None)
        if not (m_cw and y_cw) or (y_cw.get("npw") or 0) < (m_cw.get("npw") or 0):
            ytd = []

    # A segment block has to add up before it is believed. Corroboration is
    # cheap here and the alternative is bad: in the 2010-2013 releases the
    # column finder lands on the policies-in-force table, whose Agency and
    # Direct columns are counts of policies - real numbers, right labels,
    # wrong units - and nothing but arithmetic distinguishes them.
    roles = {c["role"]: c for c in segs if "role" in c}
    cw_col = roles.get("companywide")
    personal = roles.get("personal_total")
    # Everything that sits beside Personal Lines under companywide - which is
    # Commercial always, and Property only in the years it was reported apart.
    outside = [c for c in segs
               if c.get("role") in ("commercial", "property")
               and not c.get("inside_personal")]
    if cw_col and personal and cw_col.get("npw") is not None:
        total = (personal.get("npw") or 0.0) + sum(c.get("npw") or 0.0
                                                   for c in outside)
        if abs(total - cw_col["npw"]) > 1.0:
            segs = []

    # Refuse a segment block that does not agree with the headline on the same
    # page. In the 2010-2012 releases the column finder latches onto the
    # policies-in-force table, whose "Agency" and "Direct" are counts of
    # policies rather than premiums - numbers that are perfectly real, in the
    # wrong units, under the right labels. Corroboration is the only thing that
    # tells those apart from the figures wanted, so a block that fails it is
    # dropped rather than written out.
    cw = next((c for c in segs if c.get("role") == "companywide"), None)
    if head.get("npw") and cw and cw.get("npw") is not None:
        if abs(cw["npw"] - head["npw"]["ac"]) > 1.0:
            segs = []
    elif segs and not cw:
        segs = []

    return {
        "period_source": source,
        "year": period[0] if period else None,
        "month": period[1] if period else None,
        "headline": head,
        "segments": segs,
        "segments_ytd": ytd,
        "rows": len(rows),
    }


# Every US state and DC. Built as one-word names plus the multi-word ones
# spelled out, because splitting a block of text on whitespace turns "new
# jersey" into two states that do not exist and loses the one that does.
_ONE_WORD = """
alabama alaska arizona arkansas california colorado connecticut delaware
florida georgia hawaii idaho illinois indiana iowa kansas kentucky louisiana
maine maryland massachusetts michigan minnesota mississippi missouri montana
nebraska nevada ohio oklahoma oregon pennsylvania tennessee texas utah vermont
virginia washington wisconsin wyoming
""".split()

_MULTI_WORD = ("new hampshire", "new jersey", "new mexico", "new york",
               "north carolina", "north dakota", "rhode island",
               "south carolina", "south dakota", "west virginia",
               "district of columbia")

STATES = frozenset(_ONE_WORD) | frozenset(_MULTI_WORD)


def _state_name(cell: str) -> str | None:
    text = re.sub(r"\s+", " ", cell).strip().rstrip(":").lower()
    if text in ("all other", "total"):
        return text
    # Two-word states arrive as one cell; the set above holds them joined.
    joined = text.replace(".", "")
    for name in STATES:
        if name and joined == name:
            return name
    return None


def state_table(rows: list[list[str]]) -> dict[int, dict[str, float]] | None:
    """Net premiums written by state, keyed by year then row label.

    The 10-K carries five years side by side, each as an amount and a share, so
    one filing covers five years and four cover the whole period.

    Anchored on the **states**, not on the header. A 10-K is full of tables
    headed "($ in millions)" over a row of years; searching for that finds the
    first of dozens and then looks for states underneath it, where there are
    none. Searching for a run of state names instead finds exactly one place in
    the document, and the header is then whatever row of years sits above it.

    Returns None where the filing carries no such table, which is how a
    document that is not the annual-report exhibit says so.
    """
    anchor = None
    for i, cells in enumerate(rows):
        if len(cells) < 4 or not _state_name(cells[0]):
            continue
        # A run, so a single stray mention of a state cannot be mistaken for
        # the table. Three consecutive labelled rows is a table.
        run = 0
        for probe in rows[i:i + 4]:
            if probe and _state_name(probe[0]):
                run += 1
        if run >= 3:
            anchor = i
            break
    if anchor is None:
        return None

    header: list[int] | None = None
    for cells in reversed(rows[max(0, anchor - 12):anchor]):
        years = [c.strip() for c in cells
                 if re.fullmatch(r"(19|20)\d{2}", c.strip())]
        if len(years) >= 3:
            header = [int(y) for y in years]
            break
    if not header:
        return None

    out: dict[int, dict[str, float]] = {y: {} for y in header}
    seen_total = False
    for cells in rows[anchor:anchor + 60]:
        label = _state_name(cells[0])
        if not label:
            continue
        nums = [n for n in (number(c) for c in cells[1:]) if n is not None]
        # Amount and share alternate, so the amounts are every other figure.
        amounts = nums[0::2]
        if len(amounts) < len(header):
            continue
        for year, value in zip(header, amounts):
            out[year][label] = value
        if label == "total":
            seen_total = True
            break
    return out if seen_total else None
