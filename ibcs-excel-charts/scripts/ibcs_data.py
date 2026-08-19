"""Template data and structure for the IBCS proof workbook - pure Python.

No Excel, no SVG, no third-party imports. Both renderers consume this module, so
a number can only be wrong in one place.

Transcription policy
--------------------
Values are transcribed from the IBCS Institute's published template renderings in
``template-refs/``. Those renderings show *rounded* labels over decimal source
data, which matters: you cannot recover PY by subtracting the rounded absolute
variance from the rounded actual, because two roundings compound. In C03A that
error reaches 3 kEUR - enough to draw a visibly wrong grey bar.

So the transcription takes the two independently-readable series - AC (integer
labels) and the relative variance (one decimal) - and *derives* PY from them:

    PY = AC / (1 + dPY%)

which reproduces the drawn grey bars, and lets ``check_ties`` verify that the
derived absolute variance rounds back to the label IBCS actually printed. That
round-trip is the tie-out.
"""

from __future__ import annotations

import sys

from dataclasses import dataclass, field, replace
from typing import Literal, Sequence

TierKind = Literal["measure", "variance_abs", "variance_rel", "waterfall"]


# --------------------------------------------------------------------------- #
# Structure
# --------------------------------------------------------------------------- #


@dataclass(frozen=True)
class Series:
    """One scenario's values across the template's categories.

    ``values`` may contain None where a scenario does not apply to a period -
    an actual has no value in a forecast month, and vice versa. Renderers skip
    None rather than plotting zero, which would draw a bar that does not exist.
    """

    scenario: str
    values: Sequence[float | None]
    label: str | None = None

    def defined(self) -> list[tuple[int, float]]:
        return [(i, v) for i, v in enumerate(self.values) if v is not None]

    def total(self) -> float:
        return sum(v for _, v in self.defined())


@dataclass(frozen=True)
class Tier:
    """One horizontal band of a multi-tier chart.

    ``reference`` names the scenario the tier is measured against, and is what
    the axis notation encodes - solid light for PY, a double rule for PL or BU.
    A measure tier has no reference and therefore an ordinary category axis.
    """

    key: str
    label: str
    kind: TierKind
    series: Sequence[Series]
    reference: str | None = None
    height_weight: float = 1.0
    number_format: str = "{:+,.0f}"
    higher_is_better: bool = True
    printed: Sequence[float | None] | None = None
    # A table sets its columns side by side in blocks - a period block and its
    # cumulative counterpart - rather than stacking tiers down the page. The
    # block a column belongs to is a property of the column, so it lives here
    # and both renderers group by it identically.
    block: str = ""
    # The threshold past which a variance is printed in red. IBCS requires the
    # threshold to be stated wherever it is applied (UN 5.2), so the footnote
    # and the colouring must read the same number - which they can only do if
    # there is one number to read.
    red_below: float | None = None

    def series_for(self, scenario: str) -> Series | None:
        for s in self.series:
            if s.scenario == scenario:
                return s
        return None

    def label_for(self, index: int, value: float) -> str:
        """The text to put beside an element.

        Where IBCS printed a label, print theirs. Rounding a derived 5.5 gives 6,
        but the original says 5 - and a recreation that argues with the original
        about its own numbers is not a recreation. The bar is still drawn from
        the derived value, so geometry stays exact.
        """
        if self.printed is not None and self.printed[index] is not None:
            value = self.printed[index]
        text = self.number_format.format(value)
        # A signed format prints "+0.0" for a value that rounds to nothing.
        # Excel's third format section exists for exactly this case and IBCS
        # uses it - World's +0.0215% is printed "0.0%", not "+0.0%".
        if "{:+" in self.number_format:
            bare = text.replace(",", "").replace("%", "").lstrip("+-")
            if bare and float(bare) == 0:
                text = text.lstrip("+-")
        # IBCS separates thousands with a thin space, not a comma. The chart
        # renderers say this at their call sites with _thousands(); saying it
        # here as well is what lets a table go through the number format alone.
        return text.replace(",", " ")


@dataclass(frozen=True)
class Row:
    """One line of a statement drawn as a waterfall.

    A waterfall's categories are not interchangeable the way a month or a state
    is: each line either adds to a running total or subtracts from it, and some
    lines are not elements at all but the totals of the lines above them. That
    is structure, and it decides three separate things - where the bar starts,
    what shade it takes, and which direction counts as good.

    ``sign`` is +1 for a line that adds and -1 for one that subtracts. It is the
    single fact behind all three: the running total moves by ``sign * value``,
    a subtracting line is drawn in the lighter of its panel's two shades, and
    ``higher_is_better`` is simply ``sign > 0`` - which is what makes a cost
    overrun red and a tax reduction green without either being special-cased.

    ``spans`` applies to subtotals only. "zero" draws the bar from the origin -
    a result line, stating a level. An integer draws it across that many
    preceding element rows - a group line, stating a sum. Both leave the running
    total where they found it; a subtotal that also advanced it would count its
    own components twice.
    """

    label: str
    sign: int
    # "ratio" is a line the statement reports but does not add up - a margin,
    # a rate, a per-unit figure. It has to be a kind rather than a flag because
    # three separate things follow from it: it is drawn italic, its absolute
    # variance is in percentage points, and it takes no part in the walk.
    kind: Literal["element", "subtotal", "ratio"] = "element"
    prefix: str = ""          # the +/-/= the reference prints before the label
    indent: int = 0
    spans: int | str | None = None

    @property
    def higher_is_better(self) -> bool:
        return self.sign > 0


@dataclass(frozen=True)
class Annotation:
    """A highlight oval, comment reference, bracket or arrow.

    ``target`` is ``(tier_key, category_index)``; ``ref`` is the circled number
    tying the mark to a comment in the margin, or None for an unnumbered mark.
    """

    kind: Literal["oval", "bracket", "arrow", "band"]
    target: tuple[str, int]
    ref: int | None = None
    text: str | None = None


@dataclass(frozen=True)
class Comment:
    """A numbered margin comment. ``lead`` is rendered bold, as IBCS does."""

    ref: int
    lead: str
    body: str


@dataclass(frozen=True)
class TitleBlock:
    """The IBCS three-line header plus the message sentence.

    The message is the point of the chart stated as a sentence. IBCS treats it as
    the title proper - "Contribution in kEUR" is only the subject.
    """

    entity: str
    measure: str
    unit: str
    period: str
    message: str
    # A chart plotting two measures names both on the subject line, and which
    # words are bold is notation rather than decoration - the measure names are
    # bold and their units are not. Held as explicit runs because no rule
    # recovers them from a flat string, and left empty where the ordinary
    # "<measure> in <unit>" says everything. The worksheet still builds the flat
    # text from ``measure`` and ``unit``; a cell formula cannot carry a weight.
    subject: tuple[tuple[str, str], ...] = ()


def waterfall_spans(template: "Template", values: Sequence[float],
                    start: float = 0.0) -> list[tuple[float, float]]:
    """Where each bar of a waterfall starts and ends, walking one scenario.

    This lives in the data layer because it is the one part of a waterfall that
    is neither geometry nor style: it is what the numbers mean. Both renderers
    call it, so an SVG and a worksheet cannot disagree about where a bar begins
    - which they would, quickly, since the rule has three cases and only one of
    them is obvious.

        element   moves the running total by sign * value and spans the move
        subtotal, spans="zero"   spans origin to the running total: a level
        subtotal, spans=n        spans the n element rows above it: a sum

    Neither kind of subtotal advances the total. One that did would count its
    own components a second time, and the error is invisible on any row above
    the first subtotal.
    """
    spans: list[tuple[float, float]] = []
    # A statement starts at nothing; a bridge starts at the total it departs
    # from - C06F walks from the PY bar onto the AC bar.
    total = start
    for row, value in zip(template.rows, values):
        if row.kind == "ratio":
            # Reported, not added. A margin that advanced the running total
            # would have the statement adding a percentage to a revenue.
            spans.append((total, total))
        elif row.kind == "subtotal":
            if row.spans == "zero":
                spans.append((0.0, total))
            else:
                # Back up over the rows it summarises, using what they moved.
                # Counted rather than indexed, because a ratio row sitting in
                # the range occupies a position without contributing anything.
                start = total
                counted, j = 0, len(spans) - 1
                while counted < int(row.spans) and j >= 0:
                    prev = template.rows[j]
                    if prev.kind != "ratio":
                        start -= prev.sign * values[j]
                        counted += 1
                    j -= 1
                spans.append((start, total))
        else:
            moved = total + row.sign * value
            spans.append((total, moved))
            total = moved
    return spans


@dataclass(frozen=True)
class Summary:
    """A row that aggregates the categories rather than being one of them.

    Every template that shows a total needs one of these, and they turn out to
    be the same object each time: C03A's separated full-year block, C04A's USA
    row, and C06F's three scenario bars with the two variance bars beneath them.
    What varies is only how many a template has, which is why they are held as a
    sequence and why ``Template.summary`` is just the first of them.

    A summary row is drawn detached from the categories, because it is an
    aggregate of the same data rather than another member of the series. Its bar
    may stack two scenarios - in C03A, actual to date plus the hatched remainder
    of the forecast.

    ``span`` marks the rows that are a variance and nothing else: C06F draws its
    total dPY as a bar running between the AC and PY totals, with no bar of its
    own to hang it on. It names the two summary labels the bar runs between, so
    the geometry follows from the totals rather than being declared twice.
    """

    label: str
    stack: Sequence[tuple[str, float]] = ()   # [(scenario, value), ...]
    variance_abs: float | None = None
    variance_rel: float | None = None
    reference: str | None = None
    span: tuple[str, str] | None = None
    # Where the row sits relative to the categories. C03A's full-year block and
    # C04A's USA row follow them; C06F's plan and prior-year totals lead, because
    # the states are read as a decomposition of the prior year rather than as a
    # series that happens to have a total.
    before: bool = False

    @property
    def total(self) -> float:
        return sum(v for _, v in self.stack)

    @property
    def is_variance_only(self) -> bool:
        return self.span is not None


@dataclass(frozen=True)
class Segment:
    """One band of a stacked column.

    Deliberately not a ``Series``. A series is a *scenario* - what kind of
    number this is - and a stack's bands are nothing of the sort: they are
    structure categories, business areas or regions or industries, and every one
    of them is an actual. Reusing Series here would have put "Software" in a
    field whose whole job is to say AC or PL, and the notation would have had
    nothing left to carry the scenario with.
    """

    label: str
    values: Sequence[float | None]

    def at(self, index: int) -> float:
        value = self.values[index]
        return 0.0 if value is None else value


@dataclass(frozen=True)
class StructurePanel:
    """One stacked-column panel: its own categories, its own segments.

    C01 puts three of these side by side - a five-year time series, and two
    breakdowns of its last actual - and the point of the template is that they
    share a scale. Three columns of 101.9 drawn the same height is what lets a
    reader carry a total from one panel to the next; fitting each panel to its
    own range would break the only comparison the layout is making.

    ``legend_side`` is where the band names sit. IBCS wants them beside the
    bands rather than in a legend box, so the reader never has to match a colour
    to a key - and which side depends on what the panel sits next to.
    """

    key: str
    label: str
    categories: Sequence[str]
    category_scenarios: Sequence[str]
    segments: Sequence["Segment"]
    legend_side: Literal["left", "right"] = "left"
    legend_at: int = 0                # the column the band names are read off
    printed_totals: Sequence[float] | None = None

    def total(self, index: int) -> float:
        return sum(s.at(index) for s in self.segments)

    def stack(self, index: int) -> list[tuple["Segment", float, float]]:
        """Each band as (segment, base, top), walking up from the axis."""
        out, base = [], 0.0
        for segment in self.segments:
            value = segment.at(index)
            out.append((segment, base, base + value))
            base += value
        return out


@dataclass(frozen=True)
class PanelCell:
    """One cell of a small-multiples grid.

    ``key`` names the tier the cell draws, or is None for a cell left empty.
    ``row_span`` lets one panel take two rows, which is how the reference gives
    Berlin room for a variance eight times anyone else's without rescaling the
    other fourteen. ``kind`` is per cell because C13's reference panel is drawn
    as columns among fourteen panels of pins: it is the baseline the others are
    measured against, so its own variance from itself would be a row of zeros.
    """

    key: str | None
    row: int
    col: int
    row_span: int = 1
    kind: Literal["pin", "column"] = "pin"
    shaded: bool = False
    label: str | None = None      # overrides the tier's own label


@dataclass(frozen=True)
class PanelGrid:
    """One arrangement of small multiples.

    A template can carry more than one. C13's data is a single set of series,
    and the reference draws it one way while the workbook draws it another -
    the two rosters reconcile exactly, so neither is a different dataset, only
    a different view of the same one.

    ``reference`` names the tier every other panel is measured against.
    """

    rows: int
    cols: int
    cells: Sequence["PanelCell"]
    reference: str

    def drawn(self) -> list["PanelCell"]:
        return [c for c in self.cells if c.key is not None]

    def cell(self, key: str) -> "PanelCell":
        for c in self.cells:
            if c.key == key:
                return c
        raise KeyError(f"no panel cell for {key!r}")


@dataclass(frozen=True)
class TreeNode:
    """One boxed chart in a driver tree.

    A node is not a tier of one chart and not a panel of one stack - it is a
    whole small chart with its own subtitle, its own box and its own zero line,
    which happens to stand in an arithmetic relation to the others. ``key``
    names the ``Tier`` carrying its numbers, so the values live where every
    other template's values live and only the arrangement is new.

    ``scale_group`` is the rule the template exists to demonstrate. IBCS
    requires charts measuring the same thing to be drawn at the same scale, so
    that a reader carrying an impression from one box to the next is not
    misled. Boxes sharing a group share px-per-unit; boxes in different groups
    are free, because there is nothing to mislead anyone about between a
    percentage and a currency.
    """

    key: str
    column: int               # 0 = the result, rising to the base measures
    scale_group: str
    # What the box is counted in, printed after its measure name. Held per node
    # rather than on the template because a tree has no single unit - that is
    # what makes the scale groups necessary in the first place. A dimensionless
    # ratio carries none, and naming one of the two currencies it divides would
    # be wrong.
    unit: str = ""


@dataclass(frozen=True)
class TreeLink:
    """``result = left <operator> right``.

    Held as arithmetic rather than as drawing. The renderers turn a link into a
    spine, two stubs and a glyph, but the tie-outs evaluate it - which is what
    makes the picture a claim that can be false. A tree whose links were only
    lines would draw exactly the same if the numbers disagreed.
    """

    operator: Literal["x", ":"]
    left: str
    right: str
    result: str

    @property
    def glyph(self) -> str:
        """What the operator circle prints. IBCS uses the European colon."""
        return "x" if self.operator == "x" else ":"

    def apply(self, left: float, right: float) -> float:
        return left * right if self.operator == "x" else left / right


@dataclass(frozen=True)
class TreeSpec:
    """The nodes of a driver tree and the arithmetic joining them."""

    nodes: Sequence["TreeNode"]
    links: Sequence["TreeLink"]
    # Which category indices print a label. The reference labels five of C11's
    # seven years in every box, including the boxes whose columns are all
    # positive - so it is one decision applied to six charts, not collision
    # avoidance in one, and both engines have to make it identically or the six
    # category axes stop reading as the same axis.
    printed_categories: Sequence[int] = ()

    def node(self, key: str) -> "TreeNode":
        for n in self.nodes:
            if n.key == key:
                return n
        raise KeyError(f"no tree node {key!r}")

    def groups(self) -> dict[str, list[str]]:
        """Scale group -> the node keys drawn at that scale."""
        out: dict[str, list[str]] = {}
        for n in self.nodes:
            out.setdefault(n.scale_group, []).append(n.key)
        return out


@dataclass(frozen=True)
class Axis:
    """One axis of an XY chart: what it measures and how far it runs.

    A category chart gets its axis from the data - twelve months make twelve
    slots. An XY chart does not: both axes are continuous, and where they start
    and stop is an editorial decision that changes what the picture says. Push
    the top of C10's share axis from 1.25 to 2.00 and the whole portfolio slides
    into the bottom half, saying "we lead nothing" about the very same numbers.

    So the extent is declared here, as data, rather than fitted by either
    renderer. That is also what lets the two engines agree: Excel would
    otherwise choose its own round numbers and quietly disagree with the SVG.
    """

    title: str
    minimum: float
    maximum: float
    step: float
    number_format: str = "{:.2f}"
    # What the axis counts in, where it has a unit of its own. A portfolio
    # chart's axes are indices and ratios and have none; a scattergram's are
    # money and a percentage, and each states its unit under its own title
    # because no single one can head the chart.
    unit: str = ""

    def ticks(self) -> list[float]:
        out, value, n = [], self.minimum, 0
        while value <= self.maximum + self.step * 1e-6:
            out.append(round(value, 10))
            n += 1
            value = self.minimum + self.step * n
        return out

    def fraction(self, value: float) -> float:
        """Where a value sits along the axis, 0 at the minimum and 1 at the max."""
        return (value - self.minimum) / (self.maximum - self.minimum)

    def label(self, value: float) -> str:
        return self.number_format.format(value)


@dataclass(frozen=True)
class Point:
    """One plotted point in an XY chart.

    Deliberately not a ``Series`` entry. A series is one scenario read across
    shared categories, and an XY chart has no shared categories: each point
    carries its own two coordinates, and two points of the same entity sit in
    different places rather than in the same slot on different rows.

    ``size`` is the third measure where the chart has one - a bubble's area.
    It is area rather than radius, always: a radius-proportional bubble
    overstates a large value by the square, and the eye reads area.

    ``group`` names a categorical class - C09's product lines - which is
    coloured from the accent ramp rather than the scenario fills, because a
    product line is not a scenario and must not be notated as one.
    """

    entity: str
    scenario: str
    x: float
    y: float
    size: float | None = None
    group: str = ""


@dataclass(frozen=True)
class TableColumnPlan:
    """One entry in a table's left-to-right column order.

    ``gap`` says what separates this column from the one before it: nothing for
    the first, "same" for a column measured against the same reference, "group"
    for the start of a new reference. That is the only spacing rule a table
    needs, and deriving it from ``reference`` means the gutters cannot disagree
    with the header rules, which say the same thing.
    """

    kind: Literal["label", "value"]
    gap: Literal["first", "same", "group"]
    tier: "Tier | None" = None
    series_index: int = 0

    @property
    def series(self) -> "Series":
        return self.tier.series[self.series_index]

    @property
    def caption(self) -> str:
        """What the column header says, or nothing if a neighbour says it.

        A measure column falls back to its scenario name - that *is* its
        header. A variance column does not: an empty label means the column is
        the second half of a pair whose first half carries the heading, and
        falling back would head it "AC".
        """
        if self.kind == "label":
            return ""
        if self.tier.kind == "measure":
            return self.tier.label or self.series.label or self.series.scenario
        return self.tier.label

    def value(self, row: int) -> float | None:
        return self.series.values[row]


def table_column_plan(t: "Template", label_after_block: int = 1
                      ) -> list[TableColumnPlan]:
    """The left-to-right order of a table's columns, row labels included.

    Both renderers read this. The SVG turns the gaps into pixels and the
    worksheet turns them into column widths, but neither decides the order, so
    a table cannot come out with its columns in one arrangement on the page and
    another in the workbook.
    """
    blocks: dict[str, list[Tier]] = {}
    for tier in t.tiers:
        blocks.setdefault(tier.block, []).append(tier)

    plan: list[TableColumnPlan] = []
    for n, tiers in enumerate(blocks.values()):
        if n == label_after_block:
            # "first" when nothing has been placed yet: a table whose labels
            # lead starts at the margin, not one gutter in from it.
            plan.append(TableColumnPlan("label", "group" if plan else "first"))
        first, previous = True, None
        for tier in tiers:
            for index in range(len(tier.series)):
                if first:
                    # A block always starts a group. Testing the reference here
                    # would compare a measure tier's None against the sentinel
                    # None and call the first column of a block "same".
                    gap = "first" if not plan else "group"
                else:
                    gap = "same" if tier.reference == previous else "group"
                plan.append(TableColumnPlan("value", gap, tier, index))
                first, previous = False, tier.reference
    if label_after_block >= len(blocks):
        plan.append(TableColumnPlan("label", "group"))
    return plan


@dataclass(frozen=True)
class Template:
    """One IBCS template recreation."""

    id: str
    variant: str
    kind: str
    title: TitleBlock
    categories: Sequence[str]
    category_scenarios: Sequence[str]
    tiers: Sequence[Tier]
    # Columns run up from a category axis along the bottom; bars run right from
    # one down the side. IBCS treats this as a semantic choice, not a stylistic
    # one - time goes across, structure goes down - so it belongs to the template
    # rather than to the renderer.
    orientation: Literal["vertical", "horizontal"] = "vertical"
    # Set on statement templates, where a category is a line that adds or
    # subtracts rather than an interchangeable member of a series.
    rows: Sequence[Row] = field(default_factory=tuple)
    # Rows that aggregate the categories. Held as a sequence because a template
    # may have several - C06F has five - and read by both renderers, so the two
    # engines cannot disagree about what a total is.
    summary_rows: Sequence[Summary] = field(default_factory=tuple)
    # Which columns are drawn rather than printed. T01 and T02 hold the same
    # kind of figures and differ in exactly this, so it is a fact about the
    # template and both renderers read it here rather than each being told.
    panel_tiers: Sequence[str] = field(default_factory=tuple)
    # (numerator row, denominator row) for a template carrying a ratio row.
    # Held here so the worksheet builds the same quotient the page prints,
    # rather than each engine being told which two lines a margin is of.
    ratio_of: tuple[int, int] = ()
    # XY templates: the points, and the two axes they are plotted against.
    # A point is not a tier and not a category, so it gets its own field rather
    # than being forced into either.
    points: Sequence["Point"] = field(default_factory=tuple)
    axes: tuple["Axis", "Axis"] | None = None
    # (measure, unit) for the channel that carries neither axis - a bubble's
    # area. Named here because it is the chart's third measure, not decoration.
    size_legend: tuple[str, str] | None = None
    # Stacked-column panels, for the structure templates. Held separately from
    # ``tiers`` because a panel is not a band of one chart - it is a chart, and
    # C01 has three of them sharing a scale.
    structure_panels: Sequence["StructurePanel"] = field(default_factory=tuple)
    # A driver tree: several small charts and the arithmetic between them.
    # Separate from ``structure_panels`` because those share one category axis
    # and one scale; a tree's boxes share neither - each has its own zero line,
    # and the scale is shared only within a unit.
    tree: "TreeSpec | None" = None
    # Small-multiple arrangements, by name. More than one, because a single
    # set of series can be laid out more than one way and C13 is drawn both
    # the reference's way and a uniform way - the rosters reconcile exactly,
    # so neither is a different dataset.
    panel_grids: dict = field(default_factory=dict)
    annotations: Sequence[Annotation] = field(default_factory=tuple)
    comments: Sequence[Comment] = field(default_factory=tuple)
    source_ref: str = ""
    notes: Sequence[str] = field(default_factory=tuple)

    def __post_init__(self) -> None:
        keys = [t.key for t in self.tiers]
        duplicates = sorted({k for k in keys if keys.count(k) > 1})
        if duplicates:
            raise ValueError(
                f"{self.id}{self.variant} has more than one tier called "
                f"{', '.join(duplicates)}. A tier key addresses a column or "
                f"a panel, "
                f"so a repeat means one of them can never be reached."
            )

    @property
    def summary(self) -> Summary | None:
        """The first summary row, for the templates that have exactly one."""
        return self.summary_rows[0] if self.summary_rows else None

    def sheet_rows(self) -> list[tuple[str, int]]:
        """Every row a sheet needs, in the order it needs them.

        ``("category", i)`` or ``("summary", i)``. A worksheet's category axis
        has to carry the totals as well as the categories - they share it in the
        drawing, so they have to share it in the data - and this is the one
        place that decides what order they come in.
        """
        lead = [("summary", i) for i, r in enumerate(self.summary_rows) if r.before]
        trail = [("summary", i) for i, r in enumerate(self.summary_rows)
                 if not r.before]
        return lead + [("category", i) for i in range(len(self.categories))] + trail

    def summary_row(self, label: str) -> Summary:
        for row in self.summary_rows:
            if row.label == label:
                return row
        raise KeyError(f"{self.id}{self.variant} has no summary row {label!r}")

    def tier(self, key: str) -> Tier:
        for t in self.tiers:
            if t.key == key:
                return t
        raise KeyError(f"{self.id}{self.variant} has no tier {key!r}")

    def higher_is_better_at(self, index: int, tier: Tier) -> bool:
        """Which direction is good for one element of one tier.

        On a statement this is a property of the *line*, not of the tier: within
        one dPY column, revenue up is green and cost up is red. Tiers that are
        not statements keep their single flag, so nothing else changes.
        """
        if self.rows:
            return self.rows[index].higher_is_better
        return tier.higher_is_better

    def base_tier(self) -> Tier:
        """The measure tier - what the 'simple' version of this template reduces to."""
        for t in self.tiers:
            if t.kind == "measure":
                return t
        raise ValueError(f"{self.id}{self.variant} has no measure tier")


# --------------------------------------------------------------------------- #
# C03A - multi-tier column chart, Furniture Inc. contribution 2025
# --------------------------------------------------------------------------- #

MONTHS = ("Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

# Scenario per period: actual through September, forecast for the last quarter.
C03A_SCENARIOS = ("AC",) * 9 + ("FC",) * 3

# Transcribed from template-refs/C03_03A.png. These two series are the ones the
# rendering states directly; everything else here is derived from them.
C03A_MEASURE = (128, 147, 136, 139, 170, 193, 178, 189, 178, 179, 189, 211)

# September reads 3.2 on the chart, but the reference prints only one decimal, so
# what it states is the interval [3.15, 3.25). Taking 3.2 literally derives a
# variance of 5.52, which rounds to 6 where IBCS printed 5 - a disagreement
# invented by our own reading precision, not present in the source. Any value
# below 3.188 reproduces the printed label, so 3.17 is chosen: still 3.2 to the
# precision the reference states, and consistent with every label it prints.
#
# This matters more than it looks. Once the workbook is formula-driven, Excel
# derives the labels itself, so there is nowhere left to override them - the
# source data has to be right rather than corrected downstream.
C03A_VAR_REL = (11.6, 10.0, -3.4, -8.5, -2.8, 5.5, 17.3, 15.2, 3.17, -2.9, -6.2, 15.5)

# The absolute-variance labels IBCS printed. Not used to build the chart - they
# are the target that check_ties() proves the derived values reproduce.
C03A_VAR_ABS_PRINTED = (13, 13, -5, -13, -5, 10, 26, 25, 5, -5, -12, 28)


def excel_round(value: float, digits: int = 0) -> float:
    """Round half away from zero, the way Excel does.

    Python's round() goes to even, so it takes -12.5 to -12 while Excel takes it
    to -13. Any check that predicts what a cell will display has to use Excel's
    rule, or it passes on data the workbook will render differently.
    """
    factor = 10 ** digits
    scaled = value * factor
    return (int(scaled + 0.5) if scaled >= 0 else -int(-scaled + 0.5)) / factor


def _derive_py(measure: Sequence[float], var_rel: Sequence[float]) -> list[float]:
    """PY = AC / (1 + dPY%). See the transcription policy in the module docstring."""
    return [m / (1 + r / 100) for m, r in zip(measure, var_rel)]


C03A_PY = _derive_py(C03A_MEASURE, C03A_VAR_REL)
C03A_VAR_ABS = [m - p for m, p in zip(C03A_MEASURE, C03A_PY)]


def _split(values: Sequence[float | None], keep: str,
           scenarios: Sequence[str] = None) -> list[float | None]:
    """Mask a series down to the periods matching one scenario.

    ``scenarios`` defaults to C03A's for the callers that predate the
    parameter. It is worth passing explicitly: C03A turns to forecast in
    October and C07C in September, so a template that takes the default when it
    should not gets a mask that is off by a month - and an off-by-one mask
    produces a chart that renders perfectly and is wrong.
    """
    pattern = C03A_SCENARIOS if scenarios is None else scenarios
    return [v if s == keep else None for v, s in zip(values, pattern)]


C03A = Template(
    id="C03",
    variant="A",
    kind="multi-tier column chart",
    title=TitleBlock(
        entity="Furniture Inc.",
        measure="Contribution",
        unit="kEUR",
        period="2025",
        # The original render reads "prevoious"; corrected here deliberately.
        message=(
            "During the next two months the contribution will be below previous year "
            "but we expect an annual growth of 81 kEUR (+4.2%) for the full year"
        ),
    ),
    categories=MONTHS,
    category_scenarios=C03A_SCENARIOS,
    tiers=(
        Tier(
            key="var_rel",
            label="ΔPY%",
            kind="variance_rel",
            reference="PY",
            height_weight=0.85,
            number_format="{:+.1f}",
            series=(
                Series("AC", _split(C03A_VAR_REL, "AC")),
                Series("FC", _split(C03A_VAR_REL, "FC")),
            ),
        ),
        Tier(
            key="var_abs",
            label="ΔPY",
            kind="variance_abs",
            reference="PY",
            height_weight=0.85,
            number_format="{:+,.0f}",
            printed=C03A_VAR_ABS_PRINTED,
            series=(
                Series("AC", _split(C03A_VAR_ABS, "AC")),
                Series("FC", _split(C03A_VAR_ABS, "FC")),
            ),
        ),
        Tier(
            key="measure",
            label="",
            kind="measure",
            height_weight=2.6,
            number_format="{:,.0f}",
            series=(
                Series("PY", C03A_PY, label="PY"),
                Series("AC", _split(C03A_MEASURE, "AC"), label="AC"),
                Series("FC", _split(C03A_MEASURE, "FC"), label="FC"),
            ),
        ),
    ),
    summary_rows=(Summary(
        label="2025",
        stack=(("AC", 1458), ("FC", 578)),
        variance_abs=81,
        variance_rel=4.2,
        reference="PY",
    ),),
    annotations=(
        Annotation("oval", ("var_abs", -1), ref=1),
        Annotation("oval", ("measure", 11), ref=2),
    ),
    comments=(
        Comment(
            ref=1,
            lead="2025 forecast",
            body=("is 81 kEUR (+4.2%) higher than previous year. This is not including "
                  "the new product line F23 and no currency effects are expected."),
        ),
        Comment(
            ref=2,
            lead="December forecast",
            body=("of 211 kEUR is 28 kEUR (+15.5%) higher than previous year mainly "
                  "because of the additional business with Beta (30 kEUR per month)."),
        ),
    ),
    source_ref="template-refs/C03_03A.png",
    notes=(
        "AC and dPY% are transcribed; PY and dPY are derived. See check_ties().",
        "The original render carries an IBCS Institute copyright footer. Our "
        "recreation must carry our own footer, not theirs.",
    ),
)


# --------------------------------------------------------------------------- #
# Tie-outs
# --------------------------------------------------------------------------- #


class TieError(AssertionError):
    """A tie-out failed. Nothing renders until this stops being raised."""


@dataclass
class TieResult:
    name: str
    ok: bool
    detail: str

    def __str__(self) -> str:
        return f"[{'PASS' if self.ok else 'FAIL'}] {self.name}: {self.detail}"


def check_ties(template: Template = C03A, *, raise_on_fail: bool = True) -> list[TieResult]:
    """Verify a template's arithmetic against the labels IBCS actually printed.

    Rounding tolerance is deliberate, not sloppy. The reference renders integer
    labels over decimal data, so a derived variance of 5.5 can legitimately print
    as either 5 or 6. A tolerance below 1.0 on those labels would fail on correct
    data; a tolerance above it would hide a real transcription error.
    """
    results: list[TieResult] = []

    def check(name: str, ok: bool, detail: str) -> None:
        results.append(TieResult(name, ok, detail))

    if template.id in _CHECKS:
        _CHECKS[template.id](check)
        if raise_on_fail:
            failed = [r for r in results if not r.ok]
            if failed:
                raise TieError(
                    f"{len(failed)} tie-out(s) failed for {template.id}{template.variant}:\n"
                    + "\n".join(f"  {r}" for r in failed))
        return results

    if template.id != "C03":
        raise NotImplementedError(f"no tie-outs written for {template.id} yet")

    # 1. Every derived absolute variance must round to exactly the label IBCS
    #    printed. This is the strongest form of the check and the one that
    #    matters now that Excel derives the labels itself from the data - there
    #    is no longer a place to correct a disagreement downstream.
    mismatched = [
        f"{MONTHS[i]} derives {derived:+.2f} -> {excel_round(derived):+.0f}, "
        f"printed {printed:+d}"
        for i, (derived, printed) in enumerate(zip(C03A_VAR_ABS, C03A_VAR_ABS_PRINTED))
        if excel_round(derived) != printed
    ]
    check(
        "every derived dPY rounds to the printed label",
        not mismatched,
        "all 12 agree" if not mismatched else "; ".join(mismatched),
    )

    # 2. Relative variance must equal absolute variance over the base.
    worst = 0.0
    for derived, base, printed in zip(C03A_VAR_ABS, C03A_PY, C03A_VAR_REL):
        worst = max(worst, abs(derived / base * 100 - printed))
    check(
        "dPY% equals dPY over PY",
        worst < 1e-9,
        f"largest deviation {worst:.2e} percentage points",
    )

    # 3. Actual months must sum to the actual portion of the summary block.
    ac_sum = sum(v for v in _split(C03A_MEASURE, "AC") if v is not None)
    ac_block = dict(template.summary.stack)["AC"]
    check(
        "Jan-Sep actuals sum to the summary AC block",
        ac_sum == ac_block,
        f"{ac_sum:,.0f} vs {ac_block:,.0f} kEUR",
    )

    # 4. Forecast months must sum to the forecast portion, within display rounding.
    fc_sum = sum(v for v in _split(C03A_MEASURE, "FC") if v is not None)
    fc_block = dict(template.summary.stack)["FC"]
    check(
        "Oct-Dec forecasts sum to the summary FC block",
        abs(fc_sum - fc_block) <= 1.5,
        f"{fc_sum:,.0f} vs {fc_block:,.0f} kEUR "
        f"(differ by {abs(fc_sum - fc_block):.0f}; three roundings of +-0.5)",
    )

    # 5. The summary block's own variance pair must be internally consistent.
    total = sum(v for _, v in template.summary.stack)
    py_year = total / (1 + template.summary.variance_rel / 100)
    implied = total - py_year
    check(
        "summary dPY agrees with summary dPY%",
        abs(implied - template.summary.variance_abs) <= 1.5,
        f"dPY% of {template.summary.variance_rel:+.1f} implies {implied:+.1f}, "
        f"printed {template.summary.variance_abs:+.0f} kEUR",
    )

    # 6. No scenario may claim a period that belongs to another.
    measure = template.tier("measure")
    overlap = [
        MONTHS[i]
        for i in range(len(MONTHS))
        if (measure.series_for("AC").values[i] is not None)
        and (measure.series_for("FC").values[i] is not None)
    ]
    check(
        "AC and FC do not overlap",
        not overlap,
        "no period carries both" if not overlap else f"overlapping: {', '.join(overlap)}",
    )

    if raise_on_fail:
        failed = [r for r in results if not r.ok]
        if failed:
            raise TieError(
                f"{len(failed)} tie-out(s) failed for {template.id}{template.variant}:\n"
                + "\n".join(f"  {r}" for r in failed)
            )
    return results


# --------------------------------------------------------------------------- #
# C04A - multi-tier bar chart, Housing and Construction Inc. net sales 2025
# --------------------------------------------------------------------------- #
#
# The pair to C03A: same three-tier structure, but bars rather than columns, and
# measured against plan rather than prior year. That makes it the sharpest test
# of whether the abstractions generalise - anything that breaks is the shape of
# the code, not the shape of the data.
#
# Transcription differs from C03A's in a useful way. Here IBCS prints the actual
# and the *absolute* variance, so plan is the derived series: PL = AC - dPL. The
# rows are sorted by dPL descending, which is the message of this variant.

C04A_CATEGORIES = (
    "California", "Ohio", "New York", "New Hampshire", "South Carolina",
    "Hawaii", "Nevada", "New Mexico", "Maryland", "Maine", "Kansas",
    "Connecticut", "Texas", "Wisconsin", "Alaska", "Missouri", "New Jersey",
    "Virginia", "Rest of USA",
)

C04A_MEASURE = (188, 131, 211, 79, 46, 49, 23, 73, 125, 111, 45,
                50, 197, 76, 15, 59, 78, 69, 219)
C04A_VAR_ABS = (34, 31, 10, 9, 6, 6, 5, 4, 2, 2, 2,
                -3, -4, -6, -6, -29, -31, -97, 20)

# The relative-variance labels IBCS printed, in percent.
C04A_VAR_REL_PRINTED = (22, 31, 5, 13, 15, 14, 28, 6, 2, 2, 5,
                        -6, -2, -7, -29, -33, -28, -58, 10)

# Printed totals, used as an independent tie-out rather than as inputs.
C04A_TOTAL_AC, C04A_TOTAL_PL = 1845, 1889

C04A_PL = [m - v for m, v in zip(C04A_MEASURE, C04A_VAR_ABS)]
C04A_VAR_REL = [v / p * 100 for v, p in zip(C04A_VAR_ABS, C04A_PL)]

C04A = Template(
    id="C04",
    variant="A",
    kind="multi-tier bar chart",
    orientation="horizontal",
    title=TitleBlock(
        entity="Housing and Construction Inc.",
        measure="Net sales",
        unit="kUSD",
        period="2025",
        message=("Compared to plan California (+34 kUSD) and Ohio (+31 kUSD) "
                 "have the greatest absolute positive variances in net sales"),
    ),
    categories=C04A_CATEGORIES,
    category_scenarios=("AC",) * len(C04A_CATEGORIES),
    tiers=(
        Tier(
            key="measure",
            label="",
            kind="measure",
            height_weight=1.0,
            number_format="{:,.0f}",
            series=(
                Series("PL", C04A_PL, label="PL"),
                Series("AC", C04A_MEASURE, label="AC"),
            ),
        ),
        Tier(
            key="var_abs",
            label="ΔPL",
            kind="variance_abs",
            reference="PL",
            number_format="{:+,.0f}",
            printed=C04A_VAR_ABS,
            series=(Series("AC", C04A_VAR_ABS),),
        ),
        Tier(
            key="var_rel",
            label="ΔPL%",
            kind="variance_rel",
            reference="PL",
            number_format="{:+.0f}",
            printed=C04A_VAR_REL_PRINTED,
            series=(Series("AC", C04A_VAR_REL),),
        ),
    ),
    # The USA row. These are the totals IBCS printed, not sums we compute: the
    # transcribed actuals add to 1 844 against a printed 1 845, because nineteen
    # rounded values do not have to add to the rounded total. Recomputing would
    # also make dPL% meaningless - percentages do not sum.
    summary_rows=(Summary(
        label="USA",
        stack=(("AC", C04A_TOTAL_AC), ("PL", C04A_TOTAL_PL)),
        variance_abs=-45,
        variance_rel=-2,
        reference="PL",
    ),),
    annotations=(
        Annotation("oval", ("var_abs", 0)),
        Annotation("oval", ("var_abs", 1)),
        # The reference render also drops an arrow onto the first of the two,
        # pointing at the single number the message leads with.
        Annotation("arrow", ("var_abs", 0)),
    ),
    source_ref="template-refs/C04_04A.png",
    notes=(
        "AC and ΔPL are transcribed; PL and ΔPL% are derived. See check_ties().",
        "Rows are sorted by ΔPL descending - that ordering is the message.",
    ),
)


def _check_c04a(check) -> None:
    """Tie-outs for C04A.

    The strong one is the plan total. PL is derived row by row as AC - dPL, and
    IBCS separately prints a total of 1 889 that we never feed in - so if the
    nineteen derived values sum to exactly that, both transcribed series are
    right. It is the closest thing to an independent audit the source offers.
    """
    mismatched = [
        f"{C04A_CATEGORIES[i]} derives {rel:+.2f}% -> {excel_round(rel):+.0f}, "
        f"printed {printed:+d}"
        for i, (rel, printed) in enumerate(zip(C04A_VAR_REL, C04A_VAR_REL_PRINTED))
        if excel_round(rel) != printed
    ]
    check(
        "every derived dPL% rounds to the printed label",
        not mismatched,
        f"all {len(C04A_CATEGORIES)} agree" if not mismatched else "; ".join(mismatched),
    )

    pl_total = sum(C04A_PL)
    check(
        "derived PL sums to the printed plan total",
        pl_total == C04A_TOTAL_PL,
        f"{pl_total:,.0f} vs {C04A_TOTAL_PL:,.0f} kUSD (independent of our inputs)",
    )

    ac_total = sum(C04A_MEASURE)
    check(
        "transcribed AC sums to the printed actual total",
        abs(ac_total - C04A_TOTAL_AC) <= 1.5,
        f"{ac_total:,.0f} vs {C04A_TOTAL_AC:,.0f} kUSD "
        f"(differ by {abs(ac_total - C04A_TOTAL_AC)}; nineteen roundings)",
    )

    check(
        "rows are sorted by dPL descending, which is this variant's message",
        list(C04A_VAR_ABS[:-1]) == sorted(C04A_VAR_ABS[:-1], reverse=True),
        "sorted" if list(C04A_VAR_ABS[:-1]) == sorted(C04A_VAR_ABS[:-1], reverse=True)
        else "out of order (the 'Rest of USA' catch-all is excluded by design)",
    )


# --------------------------------------------------------------------------- #
# C12A - profit and loss statement as two waterfalls, Software and Service Group
# --------------------------------------------------------------------------- #
#
# Transcribed from template-refs/C12_12A-1.png, which prints every number on
# every row: PY, AC, dPY and dPY% for all twenty lines. That makes this the
# first template where nothing has to be derived in order to be drawn - and,
# more usefully, the first with a large internal check. A statement is a set of
# arithmetic identities, so the six subtotals in each of the two scenarios give
# twelve independent tie-outs. A single mistyped line breaks at least one.
#
# Both PY and AC are typed; dPY and dPY% follow.

C12A_ROWS = (
    Row("Licences", +1, prefix="+"),
    Row("Consulting", +1, prefix="+"),
    Row("Support", +1, prefix="+"),
    Row("Other revenue", +1, prefix="+"),
    Row("Sales revenue", +1, kind="subtotal", prefix="=", spans="zero"),
    Row("Other op. income", +1, prefix="+"),
    # The five lines that make up operating expenses. Indented and unprefixed in
    # the original: they are components of the line below, not of the statement.
    Row("Purchases", -1, indent=1),
    Row("Material expenses", -1, indent=1),
    Row("Personnel expenses", -1, indent=1),
    Row("Amortization", -1, indent=1),
    Row("Other op. expenses", -1, indent=1),
    Row("Operating expenses", -1, kind="subtotal", prefix="-", spans=5),
    Row("Operating result", +1, kind="subtotal", prefix="=", spans="zero"),
    Row("Investment income", +1, prefix="+"),
    Row("Financial income, net", +1, prefix="+"),
    Row("Result before tax", +1, kind="subtotal", prefix="=", spans="zero"),
    Row("Income tax", -1, prefix="-"),
    Row("Result after tax", +1, kind="subtotal", prefix="=", spans="zero"),
    Row("Profit to other investors", -1, prefix="-"),
    Row("Group result", +1, kind="subtotal", prefix="=", spans="zero"),
)

C12A_PY = (713, 72, 22, 6, 813, 45, 344, 11, 76, 56, 78, 565,
           293, 43, 73, 409, 132, 277, 89, 188)
C12A_AC = (896, 90, 10, 65, 1061, 17, 379, 54, 127, 40, 152, 752,
           326, 53, 66, 445, 111, 334, 55, 279)

# What the reference prints in the two variance columns.
C12A_VAR_ABS_PRINTED = (183, 18, -12, 59, 248, -28, 35, 43, 51, -16, 74, 187,
                        33, 10, -7, 36, -21, 57, -34, 91)
C12A_VAR_REL_PRINTED = (26, 25, -55, 983, 31, -62, 10, 391, 67, -29, 95, 33,
                        11, 23, -10, 9, -16, 21, -38, 48)

# Derived, not transcribed. dPY is exact here - both operands are integers - so
# unlike C03A there is no rounding interval to reason about.
C12A_VAR_ABS = [a - p for a, p in zip(C12A_AC, C12A_PY)]
C12A_VAR_REL = [v / p * 100 for v, p in zip(C12A_VAR_ABS, C12A_PY)]

# The two relative variances that run off any sensible scale. IBCS draws these
# with outlier markers rather than rescaling the panel to them - rule UN 5.3,
# and rule CH 4.4 on why: a big relative variance of a small absolute value is
# not important enough to distort every other row on the chart.
C12A_OUTLIERS = (3, 7)          # Other revenue +983%, Material expenses +391%

C12A = Template(
    id="C12",
    variant="A",
    kind="profit and loss waterfall",
    orientation="horizontal",
    title=TitleBlock(
        entity="Software and Service Group",
        measure="Profit and loss statement",
        unit="kEUR",
        period="AC 2025, PY and ΔPY",
        message=("Compared to 2024, the higher operating expenses (+187 kEUR) "
                 "were mainly compensated by higher license sales (+183 kEUR), "
                 "leading to a higher group result (+91 kEUR)"),
    ),
    rows=C12A_ROWS,
    categories=tuple(r.label for r in C12A_ROWS),
    # Every line is an actual or a prior actual; nothing here is planned or
    # forecast, so scenario notation is carried by the two panels rather than
    # by any per-row fill.
    category_scenarios=("AC",) * len(C12A_ROWS),
    tiers=(
        Tier(key="wf_py", label="PY", kind="waterfall", number_format="{:,.0f}",
             series=(Series("PY", C12A_PY, label="PY"),)),
        Tier(key="wf_ac", label="AC", kind="waterfall", number_format="{:,.0f}",
             series=(Series("AC", C12A_AC, label="AC"),)),
        Tier(key="var_abs", label="ΔPY", kind="variance_abs", reference="PY",
             number_format="{:+,.0f}", printed=C12A_VAR_ABS_PRINTED,
             series=(Series("AC", C12A_VAR_ABS),)),
        Tier(key="var_rel", label="ΔPY%", kind="variance_rel", reference="PY",
             number_format="{:+.0f}", printed=C12A_VAR_REL_PRINTED,
             series=(Series("AC", C12A_VAR_REL),)),
    ),
    annotations=(
        # The three numbers the message names, circled in the original.
        Annotation("oval", ("var_abs", 0)),
        Annotation("oval", ("var_abs", 11)),
        Annotation("oval", ("var_abs", 19)),
    ),
    source_ref="template-refs/C12_12A-1.png",
    notes=(
        "PY and AC are both transcribed; dPY and dPY% are derived.",
        "Impact direction is per row, not per tier: a line that subtracts from "
        "the result is better lower, so +35 on purchases is red and -21 on tax "
        "is green.",
        "Rows 4 and 8 carry relative variances of +983% and +391%, drawn with "
        "outlier markers per UN 5.3 rather than by rescaling the panel.",
    ),
)


def _check_c12a(check) -> None:
    """Tie-outs for C12A.

    A statement checks itself. Every subtotal is an assertion about the lines
    above it, so walking the waterfall and comparing each subtotal against the
    number IBCS printed on that row tests the whole transcription - twice, once
    per scenario - without needing anything the reference did not state.
    """
    for scenario, values in (("PY", C12A_PY), ("AC", C12A_AC)):
        spans = waterfall_spans(C12A, values)
        wrong = []
        for i, (row, (lo, hi)) in enumerate(zip(C12A_ROWS, spans)):
            if row.kind != "subtotal":
                continue
            got = hi if row.spans == "zero" else abs(hi - lo)
            if abs(got - values[i]) > 1e-9:
                wrong.append(f"{row.label} walks to {got:g}, printed {values[i]:g}")
        check(
            f"every {scenario} subtotal equals the lines above it",
            not wrong,
            "all 6 agree" if not wrong else "; ".join(wrong),
        )

    wrong = [f"{C12A_ROWS[i].label}: {d:+g} vs printed {p:+d}"
             for i, (d, p) in enumerate(zip(C12A_VAR_ABS, C12A_VAR_ABS_PRINTED))
             if d != p]
    check("every derived dPY equals the printed label",
          not wrong,
          "all 20 agree exactly" if not wrong else "; ".join(wrong))

    wrong = [f"{C12A_ROWS[i].label}: {d:+.1f} -> {excel_round(d):+.0f}, "
             f"printed {p:+d}"
             for i, (d, p) in enumerate(zip(C12A_VAR_REL, C12A_VAR_REL_PRINTED))
             if excel_round(d) != p]
    check("every derived dPY% rounds to the printed label",
          not wrong,
          "all 20 agree" if not wrong else "; ".join(wrong))

    # The impact direction has to fall out of the statement's own structure
    # rather than a hand-written list, or it is just an opinion typed twice.
    adverse = [i for i, r in enumerate(C12A_ROWS)
               if not r.higher_is_better and C12A_VAR_ABS[i] > 0]
    check(
        "cost increases classify as adverse without being listed as such",
        adverse == [6, 7, 8, 10, 11],
        "purchases, material, personnel and other op. expenses, and their "
        "total - five rows, all of them increases on subtracting lines",
    )

    on_scale = [i for i in C12A_OUTLIERS if abs(C12A_VAR_REL[i]) < 300]
    check("the two outlier rows really are off scale",
          not on_scale,
          f"+{C12A_VAR_REL[3]:.0f}% and +{C12A_VAR_REL[7]:.0f}% against a panel "
          f"that holds roughly 270%")


# --------------------------------------------------------------------------- #
# C06F - net sales by state, bars plus a vertical waterfall, sorted by dPY
# --------------------------------------------------------------------------- #
#
# Transcribed from template-refs/C06_06F.png. C06's four variants are the same
# chart sorted four ways, and F is the one sorted by dPY - which is why the
# ordering is checked as a tie-out rather than left as an accident of typing.
#
# The waterfall here is doing something different from C12A's. C12A walks a
# statement from zero; this one bridges two totals, starting at the PY bar and
# landing on the AC bar, with one step per state. So the value walked is the
# variance itself, signed, rather than a magnitude with a sign held beside it.
#
# Transcription follows C03A: the two independently readable series are the
# measure and the relative variance, and PY is derived as AC / (1 + dPY%). Doing
# it the other way - deriving dPY% from AC and the printed dPY - reproduces the
# reference's own labels badly, because dPY% is printed to whole points and a
# small state's percentage is very sensitive: Missouri would derive -14 where
# IBCS prints -16. Deriving in this direction reproduces all fifteen printed dPY
# labels to within 1.

C06F_STATES = ("Iowa", "Alaska", "Nebraska", "Michigan", "Minnesota", "Missouri",
               "Hawaii", "Texas", "Connecticut", "Wisconsin", "Washington",
               "New York", "California", "Illinois", "Others")

C06F_MEASURE = (179, 89, 42, 67, 55, 18, 20, 134, 64, 157, 29, 216, 257, 169, 234)

# Whole percentage points, which is all the reference prints.
C06F_VAR_REL = (33, 30, 19, 1, -5, -16, -40, -10, -23, -13, -48, -12, -16, -63, 30)

# What IBCS prints in the waterfall.
C06F_VAR_ABS_PRINTED = (44, 21, 7, 1, -3, -3, -13, -15, -19, -23, -27, -30, -48,
                        -288, 54)

# The three scenario bars and the two variance bars under them, as printed.
C06F_TOTAL_PL = 1647
C06F_TOTAL_PY = 2071
C06F_TOTAL_AC = 1728
C06F_TOTAL_VAR_ABS = -343        # AC against PY, circled in the original
C06F_TOTAL_VAR_REL = -17
C06F_TOTAL_VS_PL = 82            # AC against PL

C06F_PY = _derive_py(C06F_MEASURE, C06F_VAR_REL)
C06F_VAR_ABS = [m - p for m, p in zip(C06F_MEASURE, C06F_PY)]

# Every state is an actual measured against an earlier actual; the scenario
# notation is carried by the three summary bars, not per row.
C06F_ROWS = tuple(Row(name, +1) for name in C06F_STATES)

C06F = Template(
    id="C06",
    variant="F",
    kind="bar chart with vertical waterfall",
    orientation="horizontal",
    title=TitleBlock(
        entity="Housing and Construction Inc.",
        measure="Net sales",
        unit="kUSD",
        period="Q3 2025",
        message=("In Q3 2025, the total decrease in net sales compared to PY "
                 "was 343 kUSD"),
    ),
    rows=C06F_ROWS,
    categories=C06F_STATES,
    category_scenarios=("AC",) * len(C06F_STATES),
    tiers=(
        Tier(
            key="measure",
            label="",
            kind="measure",
            number_format="{:,.0f}",
            series=(
                Series("AC", C06F_MEASURE, label="AC"),
                Series("PY", C06F_PY, label="PY"),
            ),
        ),
        # The bridge. Its series is the variance, because that is what a step
        # of this waterfall is - the walk starts at PY and lands on AC.
        Tier(
            key="wf",
            label="ΔPY",
            kind="waterfall",
            reference="PY",
            number_format="{:+,.0f}",
            printed=C06F_VAR_ABS_PRINTED,
            series=(Series("AC", C06F_VAR_ABS),),
        ),
        Tier(
            key="var_rel",
            label="ΔPY%",
            kind="variance_rel",
            reference="PY",
            number_format="{:+.0f}",
            printed=C06F_VAR_REL,
            series=(Series("AC", C06F_VAR_REL),),
        ),
    ),
    # Five rows that are not states: the three scenario totals the states sit
    # between, and the two variance bars drawn beneath them. Both engines read
    # these, so the SVG and the worksheet cannot disagree about what the totals
    # are - which they did while these lived as module constants that only the
    # SVG knew about.
    summary_rows=(
        Summary(label="PL", stack=(("PL", C06F_TOTAL_PL),), before=True),
        Summary(label="PY", stack=(("PY", C06F_TOTAL_PY),), before=True),
        Summary(label="AC", stack=(("AC", C06F_TOTAL_AC),),
                variance_abs=C06F_TOTAL_VAR_ABS,
                variance_rel=C06F_TOTAL_VAR_REL, reference="PY"),
        # Variance-only: a bar running between two of the totals above, with no
        # bar of its own. Naming the two ends rather than a position means the
        # geometry follows the totals instead of repeating them.
        Summary(label="ΔPY", variance_abs=C06F_TOTAL_VAR_ABS,
                variance_rel=C06F_TOTAL_VAR_REL, reference="PY",
                span=("AC", "PY")),
        Summary(label="vs PL", variance_abs=C06F_TOTAL_VS_PL,
                reference="PL", span=("PL", "AC")),
    ),
    annotations=(
        # The number the message names, circled in the original, and the arrow
        # marking where the waterfall crosses from growth into decline.
        Annotation("oval", ("wf_total", 0)),
        Annotation("arrow", ("wf", 3)),
    ),
    source_ref="template-refs/C06_06F.png",
    notes=(
        "AC and dPY% are transcribed; PY and dPY are derived.",
        "Rows are sorted by dPY descending - that ordering is what makes this "
        "variant F rather than E, G or H, so it is checked.",
        "'Others' is a catch-all and sits last regardless of its variance.",
        "The left panel draws AC solid and appends the shortfall in PY's light "
        "grey, so the bar reaches PY and the grey part is what was lost.",
        "Illinois' prior year derives to 457 where the reference prints 456. "
        "The reference disagrees with itself here: its own printed 169 and -288 "
        "also give 457. Ours is kept, and the difference recorded rather than "
        "patched.",
    ),
)


def _check_c06f(check) -> None:
    """Tie-outs for C06F.

    The strong one is the bridge: fifteen derived variances have to carry the
    printed PY total onto the printed AC total. Neither total is an input to the
    derivation, so a mistyped state or percentage breaks it.
    """
    wrong = [f"{C06F_STATES[i]} derives {d:+.2f} -> {excel_round(d):+.0f}, "
             f"printed {p:+d}"
             for i, (d, p) in enumerate(zip(C06F_VAR_ABS, C06F_VAR_ABS_PRINTED))
             if abs(excel_round(d) - p) > 1]
    check("every derived dPY rounds to the printed label",
          not wrong,
          "all 15 agree within 1" if not wrong else "; ".join(wrong))

    walked = C06F_TOTAL_PY + sum(C06F_VAR_ABS)
    check(
        "the waterfall carries PY onto AC",
        abs(walked - C06F_TOTAL_AC) < 2.0,
        f"{C06F_TOTAL_PY:,} + {sum(C06F_VAR_ABS):+,.1f} = {walked:,.1f} against a "
        f"printed AC of {C06F_TOTAL_AC:,} (fifteen roundings, neither total an "
        f"input)",
    )

    total = sum(C06F_MEASURE)
    check(
        "transcribed AC sums to the printed actual total",
        abs(total - C06F_TOTAL_AC) <= 2,
        f"{total:,} vs {C06F_TOTAL_AC:,} kUSD (differ by "
        f"{abs(total - C06F_TOTAL_AC)}; fifteen roundings)",
    )

    total_py = sum(C06F_PY)
    check(
        "derived PY sums to the printed prior-year total",
        abs(total_py - C06F_TOTAL_PY) <= 3,
        f"{total_py:,.0f} vs {C06F_TOTAL_PY:,} kUSD - independent of the AC "
        f"total and of the printed variances",
    )

    implied = C06F_TOTAL_VAR_ABS / C06F_TOTAL_PY * 100
    check(
        "the summary dPY and dPY% agree",
        abs(excel_round(implied) - C06F_TOTAL_VAR_REL) < 1.0,
        f"{C06F_TOTAL_VAR_ABS} over {C06F_TOTAL_PY:,} is {implied:.1f}%, "
        f"printed {C06F_TOTAL_VAR_REL:+d}",
    )

    check(
        "AC less PL is the variance printed under the waterfall",
        abs((C06F_TOTAL_AC - C06F_TOTAL_PL) - C06F_TOTAL_VS_PL) <= 1,
        f"{C06F_TOTAL_AC:,} - {C06F_TOTAL_PL:,} = "
        f"{C06F_TOTAL_AC - C06F_TOTAL_PL}, printed +{C06F_TOTAL_VS_PL}",
    )

    # The sort order is this variant's whole identity, so it is asserted rather
    # than assumed. 'Others' is a catch-all and is excluded, as in C04A.
    ranked = list(C06F_VAR_ABS_PRINTED[:-1])
    check(
        "rows are sorted by dPY descending, which is what makes this variant F",
        ranked == sorted(ranked, reverse=True),
        "sorted" if ranked == sorted(ranked, reverse=True)
        else "out of order ('Others' is excluded by design)",
    )


# --------------------------------------------------------------------------- #
# C05X - monthly columns with a horizontal waterfall bridging plan to actual
# --------------------------------------------------------------------------- #
#
# Transcribed from template-refs/C05_05X.png, and it is the cleanest dataset in
# the library: every identity closes exactly, with no rounding interval to
# reason about anywhere. Twelve monthly figures and twelve variances give the
# plan, the actual-plus-forecast total, the split between measured and expected,
# and all twelve percentages, and each of those is checked against a number the
# reference prints separately.
#
# The waterfall is a bridge, like C06F's - it starts on the plan column and
# lands on the actual-plus-forecast column - but it runs horizontally across
# twelve months rather than down a list, and its last four steps are forecast
# and therefore hatched. That is the point of this template: the notation says
# which part of the bridge was measured and which is expected.

C05X_MEASURE = (15, 13, 17, 7, 5, 6, 11, 9, 26, 22, 18, 29)
C05X_VAR_ABS = (4, 3, 7, -3, -5, -4, -4, -5, 11, 7, 3, 10)

# Actual through August, forecast for the rest - the same split as C03A, and the
# reason the template exists.
C05X_SCENARIOS = ("AC",) * 8 + ("FC",) * 4

# What the reference prints in the pin tier, to one decimal.
C05X_VAR_REL_PRINTED = (36.4, 30.0, 70.0, -30.0, -50.0, -40.0, -26.7, -35.7,
                        73.3, 46.7, 20.0, 52.6)

# The totals the reference prints separately, and which nothing below derives
# from: they are what the tie-outs test against.
C05X_TOTAL_PY = 174           # 2024 actual
C05X_TOTAL_PL = 154           # 2025 plan
C05X_TOTAL_AC = 83            # measured part of 2025
C05X_TOTAL_FC = 95            # expected part
C05X_TOTAL_VAR_ABS = 24
C05X_TOTAL_VAR_REL = 15.6

# Derived. Plan is what the actual would have been without the variance, and the
# percentage follows from the two of them.
C05X_PL = [m - v for m, v in zip(C05X_MEASURE, C05X_VAR_ABS)]
C05X_VAR_REL = [v / p * 100 for v, p in zip(C05X_VAR_ABS, C05X_PL)]

C05X_ROWS = tuple(Row(m, +1) for m in MONTHS)


def _split_by_scenario(values, keep):
    """The half of a series that belongs to one scenario, None elsewhere."""
    return [v if s == keep else None for v, s in zip(values, C05X_SCENARIOS)]


C05X = Template(
    id="C05",
    variant="X",
    kind="column chart with horizontal waterfall",
    title=TitleBlock(
        entity="Furniture Inc.",
        measure="Net sales",
        unit="kEUR",
        period="2025",
        message=("We expect a plus of 24 kEUR (+15.6%) vs plan until end of the "
                 "year because of the positive forecast beginning in September"),
    ),
    rows=C05X_ROWS,
    categories=MONTHS,
    category_scenarios=C05X_SCENARIOS,
    tiers=(
        Tier(
            key="var_rel",
            label="ΔPL%",
            kind="variance_rel",
            reference="PL",
            number_format="{:+.1f}",
            printed=C05X_VAR_REL_PRINTED,
            series=(Series("AC", C05X_VAR_REL),),
        ),
        Tier(
            key="wf",
            label="ΔPL",
            kind="waterfall",
            reference="PL",
            number_format="{:+,.0f}",
            series=(Series("AC", C05X_VAR_ABS),),
        ),
        Tier(
            key="measure",
            label="",
            kind="measure",
            number_format="{:,.0f}",
            series=(
                Series("PL", C05X_PL, label="PL"),
                Series("AC", _split_by_scenario(C05X_MEASURE, "AC"), label="AC"),
                Series("FC", _split_by_scenario(C05X_MEASURE, "FC"), label="FC"),
            ),
        ),
    ),
    summary_rows=(
        Summary(label="2024 AC", stack=(("PY", C05X_TOTAL_PY),), before=True),
        Summary(label="2025 PL", stack=(("PL", C05X_TOTAL_PL),), before=True),
        # The closing column stacks what was measured under what is expected,
        # which is the only place in the template where the two are added up.
        Summary(label="2025 AC+FC",
                stack=(("AC", C05X_TOTAL_AC), ("FC", C05X_TOTAL_FC)),
                variance_abs=C05X_TOTAL_VAR_ABS,
                variance_rel=C05X_TOTAL_VAR_REL, reference="PL"),
        # The closing column is compared with both of the columns it is drawn
        # beside, so it carries two variance bars rather than one. Each names
        # the two totals it runs between, so neither can be drawn against the
        # wrong one.
        Summary(label="vs PL", variance_abs=C05X_TOTAL_VAR_ABS, reference="PL",
                span=("2025 PL", "2025 AC+FC")),
        Summary(label="vs PY",
                variance_abs=(C05X_TOTAL_AC + C05X_TOTAL_FC) - C05X_TOTAL_PY,
                reference="PY", span=("2024 AC", "2025 AC+FC")),
    ),
    annotations=(
        # The two numbers the message names.
        Annotation("oval", ("var_rel", -1)),
        Annotation("oval", ("wf", -1)),
    ),
    source_ref="template-refs/C05_05X.png",
    notes=(
        "The measure and dPL are transcribed; PL and dPL% are derived.",
        "Nothing here needs a rounding tolerance - every identity closes exactly.",
        "September is where measured becomes expected, and the notation carries "
        "it: hatched columns, hatched waterfall steps and hatched pin heads.",
    ),
)


def _check_c05x(check) -> None:
    """Tie-outs for C05X.

    Unusually, all of these are exact. Six printed totals are tested against
    numbers derived from the twelve monthly figures, and none of the six is an
    input to the derivation - so any single mistyped month breaks at least two.
    """
    total = sum(C05X_MEASURE)
    printed_total = C05X_TOTAL_AC + C05X_TOTAL_FC
    check("the monthly figures sum to the closing column",
          total == printed_total,
          f"{total} vs {C05X_TOTAL_AC} + {C05X_TOTAL_FC} = {printed_total} kEUR")

    measured = sum(v for v, s in zip(C05X_MEASURE, C05X_SCENARIOS) if s == "AC")
    expected = sum(v for v, s in zip(C05X_MEASURE, C05X_SCENARIOS) if s == "FC")
    check("the measured and expected halves match their printed blocks",
          measured == C05X_TOTAL_AC and expected == C05X_TOTAL_FC,
          f"Jan-Aug {measured} (printed {C05X_TOTAL_AC}), "
          f"Sep-Dec {expected} (printed {C05X_TOTAL_FC})")

    check("the monthly variances sum to the printed total variance",
          sum(C05X_VAR_ABS) == C05X_TOTAL_VAR_ABS,
          f"{sum(C05X_VAR_ABS):+d} vs {C05X_TOTAL_VAR_ABS:+d} kEUR")

    check("derived plan sums to the printed plan column",
          sum(C05X_PL) == C05X_TOTAL_PL,
          f"{sum(C05X_PL)} vs {C05X_TOTAL_PL} kEUR - derived from the months, "
          f"not read off the column")

    check("the bridge carries plan onto actual plus forecast",
          C05X_TOTAL_PL + sum(C05X_VAR_ABS) == printed_total,
          f"{C05X_TOTAL_PL} {sum(C05X_VAR_ABS):+d} = "
          f"{C05X_TOTAL_PL + sum(C05X_VAR_ABS)} vs a printed {printed_total}")

    wrong = [f"{MONTHS[i]} derives {d:.1f}, printed {p:+.1f}"
             for i, (d, p) in enumerate(zip(C05X_VAR_REL, C05X_VAR_REL_PRINTED))
             if abs(d - p) > 0.05]
    check("every derived dPL% matches the printed label to one decimal",
          not wrong,
          "all 12 agree" if not wrong else "; ".join(wrong))

    implied = sum(C05X_VAR_ABS) / sum(C05X_PL) * 100
    check("the summary dPL% follows from the summary dPL",
          abs(implied - C05X_TOTAL_VAR_REL) < 0.05,
          f"{sum(C05X_VAR_ABS)} over {sum(C05X_PL)} is {implied:.2f}%, "
          f"printed {C05X_TOTAL_VAR_REL:+.1f}")

    check("actual and forecast do not overlap",
          not any(a is not None and f is not None
                  for a, f in zip(_split_by_scenario(C05X_MEASURE, "AC"),
                                  _split_by_scenario(C05X_MEASURE, "FC"))),
          "no month carries both")


# --------------------------------------------------------------------------- #
# T01B - hierarchical table, Pharmaceutical Inc. profit after tax, November 2026
# --------------------------------------------------------------------------- #
#
# The first table, and a different shape of problem from every chart before it.
# A chart states a few numbers and draws the rest; a table states all of them,
# so there is nothing to derive and nowhere for a rounding to hide. What the
# recreation has to get right is structure and notation, not geometry.
#
# Two column blocks sit side by side with the row labels between them: the
# month, and the year to date. IBCS marks the cumulative block by prefixing the
# period name with an underscore - "_November" - rather than by writing "YTD"
# (UN 4.2; a trailing underscore would mean year-to-go, a leading tilde a moving
# annual total, which is what C07 labels MAT).
#
# Every value here is transcribed. The variances are all derived, because they
# are exact: PY, PL and AC are integers, so AC - PY has no rounding interval of
# the kind that made C03A's transcription delicate. That leaves two things worth
# checking, and both are checked below - that the subtotals equal the rows they
# summarise, and that each derived percentage rounds back to the label IBCS
# printed. The first tests all 120 transcribed numbers; the second tests the
# derivation, and would also expose hidden decimals in the source, which the
# cumulative block's "+431.0" hints at but does not turn out to have.

# The row hierarchy, shared with T02A - same countries, same three groups. Every
# country adds, so a group subtotal spans the rows above it and the World row is
# simply the running total, exactly as a waterfall's "= result" line is.
TABLE_COUNTRY_ROWS = (
    Row("Austria", +1),
    Row("Belgium", +1),
    Row("France", +1),
    Row("Germany", +1),
    Row("Poland", +1),
    Row("Sweden", +1),
    Row("Switzerland", +1),
    Row("Other", +1),
    Row("Europe", +1, kind="subtotal", spans=8),
    Row("Brazil", +1),
    Row("Canada", +1),
    Row("USA", +1),
    Row("Other", +1),
    Row("Americas", +1, kind="subtotal", spans=4),
    Row("Australia", +1),
    Row("China", +1),
    Row("Japan", +1),
    Row("Other", +1),
    Row("Rest of World", +1, kind="subtotal", spans=4),
    Row("World", +1, kind="subtotal", spans="zero"),
)

# November. Subtotal rows carry the printed subtotal; check_ties proves each one
# equals the rows above it, which is what makes the transcription self-testing.
T01B_NOV_PY = (500, 56, 140, 345, 78, 77, 61, 502, 1759,
               119, 65, 346, 438, 968,
               54, 266, 9, 243, 572, 3299)
T01B_NOV_PL = (590, 72, 149, 279, 91, 81, 70, 498, 1830,
               109, 71, 326, 401, 907,
               66, 204, 12, 323, 605, 3342)
T01B_NOV_AC = (559, 58, 134, 260, 86, 86, 66, 545, 1794,
               121, 59, 311, 399, 890,
               62, 231, 11, 266, 570, 3254)

# _November - the year to date.
T01B_YTD_PY = (5078, 531, 1290, 3124, 816, 809, 604, 5602, 17854,
               1205, 629, 3406, 4166, 9406,
               517, 2107, 67, 2531, 5222, 32482)
T01B_YTD_PL = (5611, 529, 1488, 2815, 818, 722, 582, 6022, 18587,
               1254, 656, 3124, 4219, 9253,
               609, 1925, 87, 2099, 4720, 32560)
T01B_YTD_AC = (5509, 484, 1354, 2850, 854, 764, 678, 5441, 17934,
               1314, 718, 3239, 4008, 9279,
               588, 2399, 144, 2145, 5276, 32489)

# The percentage labels IBCS printed. Not inputs - the targets the derived
# values have to reproduce. The absolute variance columns are deliberately not
# transcribed: AC - PY over integers is exact, so a printed absolute variance
# would test nothing that its own percentage does not test better.
T01B_NOV_DPY_REL_PRINTED = (11.8, 3.6, -4.3, -24.6, 10.3, 11.7, 8.2, 8.6, 2.0,
                            1.7, -9.2, -10.1, -8.9, -8.1,
                            14.8, -13.2, 22.2, 9.5, -0.3, -1.4)
T01B_NOV_DPL_REL_PRINTED = (-5.3, -19.4, -10.1, -6.8, -5.5, 6.2, -5.7, 9.4, -2.0,
                            11.0, -16.9, -4.6, -0.5, -1.9,
                            -6.1, 13.2, -8.3, -17.6, -5.8, -2.6)
T01B_YTD_DPY_REL_PRINTED = (8.5, -8.9, 5.0, -8.8, 4.7, -5.6, 12.3, -2.9, 0.4,
                            9.0, 14.1, -4.9, -3.8, -1.4,
                            13.7, 13.9, 114.9, -15.3, 1.0, 0.0)
T01B_YTD_DPL_REL_PRINTED = (-1.8, -8.5, -9.0, 1.2, 4.4, 5.8, 16.5, -9.6, -3.5,
                            4.8, 9.5, 3.7, -5.0, 0.3,
                            -3.4, 24.6, 65.5, 2.2, 11.8, -0.2)


def _variances(actual, reference):
    """The absolute and relative variance of one column against another.

    A None on either side gives a None: a row a scenario does not apply to has
    no variance, and drawing a zero there would state something the source does
    not.
    """
    absolute = [None if a is None or r is None else a - r
                for a, r in zip(actual, reference)]
    relative = [None if v is None or not r else v / r * 100
                for v, r in zip(absolute, reference)]
    return absolute, relative


T01B_NOV_DPY, T01B_NOV_DPY_REL = _variances(T01B_NOV_AC, T01B_NOV_PY)
T01B_NOV_DPL, T01B_NOV_DPL_REL = _variances(T01B_NOV_AC, T01B_NOV_PL)
T01B_YTD_DPY, T01B_YTD_DPY_REL = _variances(T01B_YTD_AC, T01B_YTD_PY)
T01B_YTD_DPL, T01B_YTD_DPL_REL = _variances(T01B_YTD_AC, T01B_YTD_PL)


def _table_columns(block, py, pl, ac, dpy, dpy_rel, dpy_rel_printed,
                   dpl, dpl_rel, dpl_rel_printed, *, abs_format, red_below):
    """One block of a variance table: three scenarios and two variance pairs.

    A column of a table is the same object as a tier of a chart - a keyed band
    of values measured against a named scenario - so it is held as one, and the
    notation follows from ``reference`` in both engines without being restated.
    """
    # The cumulative block needs its own key prefix. "_November".strip("_")
    # gives "november", the same as the month block - which silently produced
    # two sets of tiers with identical keys, and a worksheet that mapped
    # fourteen columns onto seven.
    tag = ("ytd_" if block.startswith("_") else "") + block.strip("_").lower()
    return (
        Tier(key=f"m_{tag}", label="", kind="measure", block=block,
             number_format="{:,.0f}",
             series=(Series("PY", py, label="PY"),
                     Series("PL", pl, label="PL"),
                     Series("AC", ac, label="AC"))),
        Tier(key=f"dpy_{tag}", label="ΔPY", kind="variance_abs", block=block,
             reference="PY", number_format=abs_format, red_below=red_below,
             series=(Series("AC", dpy),)),
        Tier(key=f"dpyp_{tag}", label="ΔPY%", kind="variance_rel", block=block,
             reference="PY", number_format="{:+,.1f}%", red_below=-10.0,
             printed=dpy_rel_printed, series=(Series("AC", dpy_rel),)),
        Tier(key=f"dpl_{tag}", label="ΔPL", kind="variance_abs", block=block,
             reference="PL", number_format="{:+,.0f}", red_below=red_below,
             series=(Series("AC", dpl),)),
        Tier(key=f"dplp_{tag}", label="ΔPL%", kind="variance_rel", block=block,
             reference="PL", number_format="{:+,.1f}%", red_below=-10.0,
             printed=dpl_rel_printed, series=(Series("AC", dpl_rel),)),
    )


T01B = Template(
    id="T01",
    variant="B",
    kind="hierarchical table",
    title=TitleBlock(
        entity="Pharmaceutical Inc.",
        measure="Profit after tax",
        unit="kEUR",
        period="November 2026",
        message="",
    ),
    rows=TABLE_COUNTRY_ROWS,
    categories=tuple(r.label for r in TABLE_COUNTRY_ROWS),
    category_scenarios=("AC",) * len(TABLE_COUNTRY_ROWS),
    tiers=(
        *_table_columns(
            "November",
            T01B_NOV_PY, T01B_NOV_PL, T01B_NOV_AC,
            T01B_NOV_DPY, T01B_NOV_DPY_REL, T01B_NOV_DPY_REL_PRINTED,
            T01B_NOV_DPL, T01B_NOV_DPL_REL, T01B_NOV_DPL_REL_PRINTED,
            abs_format="{:+,.0f}", red_below=-20.0),
        *_table_columns(
            "_November",
            T01B_YTD_PY, T01B_YTD_PL, T01B_YTD_AC,
            T01B_YTD_DPY, T01B_YTD_DPY_REL, T01B_YTD_DPY_REL_PRINTED,
            T01B_YTD_DPL, T01B_YTD_DPL_REL, T01B_YTD_DPL_REL_PRINTED,
            # The cumulative block prints its dPY to one decimal where every
            # other variance column is an integer. Reproduced as printed rather
            # than regularised: it is the source that is inconsistent, and a
            # recreation that tidies its original is no longer a recreation.
            abs_format="{:+,.1f}", red_below=-220.0),
    ),
    source_ref="template-refs/T01_T01B.png",
    notes=(
        "Every value is transcribed; all eight variance columns are derived.",
        "The second block is the year to date, marked by the underscore prefix "
        "of UN 4.2 rather than by the word YTD.",
        "Red is a stated threshold, not a judgement: -20 and -10% on the month, "
        "-220 and -10% on the year to date, printed in the footnote.",
    ),
)


def _check_t01b(check) -> None:
    """Tie-outs for T01B.

    A hierarchical table checks itself the way a statement does. Each of the
    three group subtotals asserts the rows above it and the World row asserts
    all sixteen, in three scenarios across two blocks - six walks that between
    them touch every one of the 120 transcribed numbers. A single mistyped
    country breaks its group and the World row, and nothing else, which is what
    makes the failure legible.
    """
    blocks = (
        ("November", T01B_NOV_PY, T01B_NOV_PL, T01B_NOV_AC),
        ("_November", T01B_YTD_PY, T01B_YTD_PL, T01B_YTD_AC),
    )
    for block, py, pl, ac in blocks:
        for scenario, values in (("PY", py), ("PL", pl), ("AC", ac)):
            spans = waterfall_spans(T01B, values)
            wrong = []
            for i, (row, (lo, hi)) in enumerate(zip(TABLE_COUNTRY_ROWS, spans)):
                if row.kind != "subtotal":
                    continue
                got = hi if row.spans == "zero" else abs(hi - lo)
                if abs(got - values[i]) > 1e-9:
                    wrong.append(f"{row.label} sums to {got:g}, printed {values[i]:g}")
            check(
                f"{block} {scenario}: every subtotal equals the rows above it",
                not wrong,
                "all 4 agree exactly" if not wrong else "; ".join(wrong),
            )

    percentages = (
        ("November dPY%", T01B_NOV_DPY_REL, T01B_NOV_DPY_REL_PRINTED),
        ("November dPL%", T01B_NOV_DPL_REL, T01B_NOV_DPL_REL_PRINTED),
        ("_November dPY%", T01B_YTD_DPY_REL, T01B_YTD_DPY_REL_PRINTED),
        ("_November dPL%", T01B_YTD_DPL_REL, T01B_YTD_DPL_REL_PRINTED),
    )
    for name, derived, printed in percentages:
        wrong = [
            f"{TABLE_COUNTRY_ROWS[i].label} derives {d:+.3f}, printed {p:+.1f}"
            for i, (d, p) in enumerate(zip(derived, printed))
            if abs(excel_round(d, 1) - p) > 1e-9
        ]
        check(
            f"every derived {name} rounds to the printed label",
            not wrong,
            "all 20 agree" if not wrong else "; ".join(wrong),
        )

    # The threshold is the whole point of the red: it has to be a property of
    # the column that the footnote can read, not a rule applied by eye. This
    # asserts the two the source states, so a renderer that invents its own
    # fails here rather than quietly disagreeing with the printed footnote.
    thresholds = {(t.block, t.label): t.red_below
                  for t in T01B.tiers if t.kind != "measure"}
    check(
        "each variance column states the threshold its red is applied at",
        thresholds == {("November", "ΔPY"): -20.0, ("November", "ΔPY%"): -10.0,
                       ("November", "ΔPL"): -20.0, ("November", "ΔPL%"): -10.0,
                       ("_November", "ΔPY"): -220.0, ("_November", "ΔPY%"): -10.0,
                       ("_November", "ΔPL"): -220.0, ("_November", "ΔPL%"): -10.0},
        "-20/-10% on the month and -220/-10% on the year to date, as printed",
    )

    # Which cells the reference actually prints in red, read off the rendering.
    # Derived membership rather than a transcribed list: the test is that the
    # stated rule reproduces the source's own choices.
    #
    # Keyed by row *index*, not row label. Three rows are called "Other", so a
    # label-keyed set collapses them - and would then pass while the renderer
    # reddened the wrong continent's Other.
    red = {(t.block, t.label): frozenset(
               i for i, v in enumerate(t.series[0].values) if v < t.red_below)
           for t in T01B.tiers if t.red_below is not None}
    expected = {
        ("November", "ΔPY"): frozenset({3, 11, 12, 13, 15, 19}),
        ("November", "ΔPY%"): frozenset({3, 11, 15}),
        ("November", "ΔPL"): frozenset({0, 8, 17, 18, 19}),
        ("November", "ΔPL%"): frozenset({1, 2, 10, 17}),
        ("_November", "ΔPY"): frozenset({3, 17}),
        ("_November", "ΔPY%"): frozenset({17}),
        ("_November", "ΔPL"): frozenset({7, 8}),
        # No cumulative ΔPL% reaches -10%; the nearest is Other at -9.6%.
        ("_November", "ΔPL%"): frozenset(),
    }
    wrong = [f"{block} {col}: ours {sorted(red[block, col])}, "
             f"reference {sorted(want)}"
             for (block, col), want in expected.items()
             if red[block, col] != want]
    check(
        "the stated threshold reproduces the cells the reference prints in red",
        not wrong,
        f"{sum(len(v) for v in red.values())} cells, matching the rendering"
        if not wrong else "; ".join(wrong),
    )


# --------------------------------------------------------------------------- #
# T02A - the same table with its variance columns drawn, Pharmaceutical Inc.
# --------------------------------------------------------------------------- #
#
# T01B and T02A are the same report told two ways: one prints the variance as a
# number, the other draws it. So the rows are the same rows, the month block is
# the same month block - Austria 590 and 559 in both, Europe 1 830 and 1 794,
# World 3 342 and 3 254, all identical - and the only structural difference is
# that there is no prior year here, because a drawn panel can carry one
# comparison legibly and not two.
#
# The cumulative block is *not* shared, which is worth stating plainly because
# it would be easy to assume otherwise. T02A's Other in Europe is 5 899 against
# T01B's 5 441, so its Europe reads 18 392 against 17 934 and its World 32 947
# against 32 489. Two published examples of the same report disagree; both are
# transcribed as printed rather than reconciled, and check_ties proves each is
# internally consistent on its own terms.
#
# There is no threshold footnote here. The colour is carried by the bars, and a
# bar is coloured by impact rather than by crossing a stated line, so nothing
# needs stating.

T02A_NOV_PL = T01B_NOV_PL
T02A_NOV_AC = T01B_NOV_AC

T02A_YTD_PL = (5611, 529, 1488, 2815, 818, 722, 582, 6022, 18587,
               1254, 656, 3124, 4219, 9253,
               609, 2311, 139, 2099, 5158, 32998)
T02A_YTD_AC = (5509, 484, 1354, 2850, 854, 764, 678, 5899, 18392,
               1314, 718, 3239, 4008, 9279,
               588, 2399, 144, 2145, 5276, 32947)

T02A_NOV_DPL_REL_PRINTED = T01B_NOV_DPL_REL_PRINTED
T02A_YTD_DPL_REL_PRINTED = (-1.8, -8.5, -9.0, 1.2, 4.4, 5.8, 16.5, -2.0, -1.0,
                            4.8, 9.5, 3.7, -5.0, 0.3,
                            -3.4, 3.8, 3.6, 2.2, 2.3, -0.2)

T02A_NOV_DPL, T02A_NOV_DPL_REL = _variances(T02A_NOV_AC, T02A_NOV_PL)
T02A_YTD_DPL, T02A_YTD_DPL_REL = _variances(T02A_YTD_AC, T02A_YTD_PL)


def _panel_columns(block, pl, ac, dpl, dpl_rel, dpl_rel_printed):
    """One block of an integrated-bar table: two scenarios and one drawn pair."""
    tag = ("ytd_" if block.startswith("_") else "") + block.strip("_").lower()
    return (
        Tier(key=f"m_{tag}", label="", kind="measure", block=block,
             number_format="{:,.0f}",
             series=(Series("PL", pl, label="PL"),
                     Series("AC", ac, label="AC"))),
        Tier(key=f"dpl_{tag}", label="ΔPL", kind="variance_abs", block=block,
             reference="PL", number_format="{:+,.0f}",
             series=(Series("AC", dpl),)),
        Tier(key=f"dplp_{tag}", label="ΔPL%", kind="variance_rel", block=block,
             reference="PL", number_format="{:+,.1f}",
             printed=dpl_rel_printed, series=(Series("AC", dpl_rel),)),
    )


T02A = Template(
    id="T02",
    variant="A",
    kind="table with integrated bars",
    title=TitleBlock(
        entity="Pharmaceutical Inc.",
        measure="Profit after tax",
        unit="kEUR",
        period="November 2026",
        message="",
    ),
    rows=TABLE_COUNTRY_ROWS,
    categories=tuple(r.label for r in TABLE_COUNTRY_ROWS),
    category_scenarios=("AC",) * len(TABLE_COUNTRY_ROWS),
    tiers=(
        *_panel_columns("November", T02A_NOV_PL, T02A_NOV_AC,
                        T02A_NOV_DPL, T02A_NOV_DPL_REL,
                        T02A_NOV_DPL_REL_PRINTED),
        *_panel_columns("_November", T02A_YTD_PL, T02A_YTD_AC,
                        T02A_YTD_DPL, T02A_YTD_DPL_REL,
                        T02A_YTD_DPL_REL_PRINTED),
    ),
    # Which columns are drawn rather than printed. This is what separates T02
    # from T01 - the same figures, stated as bars - so it belongs to the
    # template, and both renderers read it rather than each being told.
    panel_tiers=("dpl_november", "dplp_november",
                 "dpl_ytd_november", "dplp_ytd_november"),
    source_ref="template-refs/T02_T02A.png",
    notes=(
        "Shares T01B's rows and its whole November block; the cumulative block "
        "genuinely differs and is transcribed separately.",
        "No threshold footnote: a drawn variance is coloured by impact, so "
        "there is no stated line to cross.",
    ),
)


def _check_t02a(check) -> None:
    """Tie-outs for T02A.

    The same six subtotal walks as T01B - four fewer, since there is no prior
    year - plus the two percentage round-trips. The interesting extra check is
    the last one: that the shared month block really is shared, which is the
    claim the transcription rests on and the one that would be embarrassing to
    assume.
    """
    for block, pl, ac in (("November", T02A_NOV_PL, T02A_NOV_AC),
                          ("_November", T02A_YTD_PL, T02A_YTD_AC)):
        for scenario, values in (("PL", pl), ("AC", ac)):
            spans = waterfall_spans(T02A, values)
            wrong = []
            for i, (row, (lo, hi)) in enumerate(zip(TABLE_COUNTRY_ROWS, spans)):
                if row.kind != "subtotal":
                    continue
                got = hi if row.spans == "zero" else abs(hi - lo)
                if abs(got - values[i]) > 1e-9:
                    wrong.append(f"{row.label} sums to {got:g}, printed {values[i]:g}")
            check(f"{block} {scenario}: every subtotal equals the rows above it",
                  not wrong,
                  "all 4 agree exactly" if not wrong else "; ".join(wrong))

    for name, derived, printed in (
            ("November dPL%", T02A_NOV_DPL_REL, T02A_NOV_DPL_REL_PRINTED),
            ("_November dPL%", T02A_YTD_DPL_REL, T02A_YTD_DPL_REL_PRINTED)):
        wrong = [f"{TABLE_COUNTRY_ROWS[i].label} derives {d:+.3f}, printed {p:+.1f}"
                 for i, (d, p) in enumerate(zip(derived, printed))
                 if abs(excel_round(d, 1) - p) > 1e-9]
        check(f"every derived {name} rounds to the printed label",
              not wrong, "all 20 agree" if not wrong else "; ".join(wrong))

    check(
        "the month block is the one T01B prints, value for value",
        T02A_NOV_PL == T01B_NOV_PL and T02A_NOV_AC == T01B_NOV_AC,
        "40 figures shared with T01B, transcribed once",
    )

    # And the cumulative block is not. Asserted rather than noted, so that a
    # later tidy-up that "fixes" one to match the other fails here.
    pl_differs = [TABLE_COUNTRY_ROWS[i].label
                  for i, (a, b) in enumerate(zip(T02A_YTD_PL, T01B_YTD_PL)) if a != b]
    ac_differs = [TABLE_COUNTRY_ROWS[i].label
                  for i, (a, b) in enumerate(zip(T02A_YTD_AC, T01B_YTD_AC)) if a != b]
    check(
        "the cumulative block differs from T01B, as the two sources do",
        (pl_differs == ["China", "Japan", "Rest of World", "World"]
         and ac_differs == ["Other", "Europe", "World"]),
        "the plans differ on China and Japan and the actuals on Europe's "
        "Other, each carrying into its own subtotals and the World row - "
        "seven rows in all, and no others",
    )


# --------------------------------------------------------------------------- #
# T03A - measure rows: a full P&L against both prior year and plan
# --------------------------------------------------------------------------- #
#
# The same grid as T01B over a statement instead of a hierarchy of countries, so
# the rows are C12A's rows and the columns are T01B's columns. Two things are new.
#
# **A ratio row.** Gross margin is operating result over sales revenue - a line
# the statement reports but does not add up. It is drawn italic and indented, its
# figures are percentages rather than kEUR, and its absolute variance is in
# percentage *points*, which is why IBCS prints "+8.6%p" rather than "+8.6".
# Structurally it must sit outside the walk: a row that took part in the running
# total would have the statement adding a margin to a revenue.
#
# **One decimal over more precise data.** The reference prints to a tenth and
# computes on the full figures, so walking the printed lines lands 0.1 away on
# some subtotals - once per scenario, then carried down. That is the C03A lesson
# again and the tie-outs below state it rather than hiding it behind a tolerance.
#
# And one outright error in the source, which is a first for this project; see
# T03A_SOURCE_ERROR below.

T03A_ROWS = (
    Row("Licences", +1, prefix="+"),
    Row("Consulting", +1, prefix="+"),
    Row("Maintenance", +1, prefix="+"),
    Row("Other revenue", +1, prefix="+"),
    Row("Sales revenue", +1, kind="subtotal", prefix="=", spans="zero"),
    Row("Other op. income", +1, prefix="+"),
    Row("Purchases", -1, indent=1),
    Row("Material expenses", -1, indent=1),
    Row("Personnel expenses", -1, indent=1),
    Row("Amortization", -1, indent=1),
    Row("Other op. expenses", -1, indent=1),
    Row("Operating expenses", -1, kind="subtotal", prefix="-", spans=5),
    Row("Operating result", +1, kind="subtotal", prefix="=", spans="zero"),
    # Not indented: it sits at the label position like every result line. What
    # makes it look indented is having no +/-/= prefix, which is the point - it
    # is not a line of the statement, it is a ratio of two of them.
    Row("Gross margin", +1, kind="ratio"),
    Row("Investment income", +1, prefix="+"),
    Row("Financial income, net", +1, prefix="+"),
    Row("Result before tax", +1, kind="subtotal", prefix="=", spans="zero"),
    Row("Income tax", -1, prefix="-"),
    Row("Result after tax", +1, kind="subtotal", prefix="=", spans="zero"),
    Row("Profit to other investors", -1, prefix="-"),
    Row("Group result", +1, kind="subtotal", prefix="=", spans="zero"),
)

T03A_MARGIN = 13          # the ratio row
T03A_RESULT, T03A_REVENUE = 12, 4      # the two rows it is the ratio of

# None marks the ratio row: it is derived from two rows of the same column, so
# typing it would be typing a total, which is the one thing this project does
# not do.
T03A_PY_LINES = (713.3, 72.0, 22.0, 15.0, 822.3, 45.0,
                 344.2, 11.0, 76.0, 56.0, 78.0, 565.2, 302.1, None,
                 43.0, 73.0, 418.2, 132.1, 286.1, 50.0, 236.1)
T03A_PL_LINES = (749.0, 75.6, 22.0, 25.0, 871.6, 35.0,
                 361.4, 11.6, 79.8, 56.0, 76.0, 584.8, 321.9, None,
                 39.0, 66.0, 426.9, 115.0, 311.9, 52.0, 259.9)
T03A_AC_LINES = (896.4, 90.0, 20.4, 24.4, 1031.2, 38.8,
                 379.2, 15.8, 87.6, 44.4, 75.6, 602.6, 467.5, None,
                 53.0, 69.6, 590.1, 111.1, 479.0, 56.7, 422.3)


def _with_margin(values):
    """Fill the ratio row from the two rows it is the ratio of."""
    out = list(values)
    out[T03A_MARGIN] = out[T03A_RESULT] / out[T03A_REVENUE] * 100
    return tuple(out)


T03A_PY = _with_margin(T03A_PY_LINES)
T03A_PL = _with_margin(T03A_PL_LINES)
T03A_AC = _with_margin(T03A_AC_LINES)

# The printed variance labels, all four columns. Transcribed here - unlike T01B,
# where the absolute columns were exact and therefore untestable - because at one
# decimal they are not exact, and the interesting question is how far off a
# derived value is allowed to be.
T03A_DPY_PRINTED = (183.1, 18.0, -1.6, 9.4, 208.9, -6.2,
                    35.0, 4.8, 11.6, -11.6, -2.4, 37.3, 165.3, 8.6,
                    10.0, -3.4, 171.9, -21.0, 192.9, 6.7, 162.5)
T03A_DPY_REL_PRINTED = (26, 25, -7, 63, 25, -14,
                        10, 44, 15, -21, -3, 7, 55, 23,
                        23, -5, 41, -16, 67, 13, 69)
T03A_DPL_PRINTED = (147.4, 14.4, -1.6, -0.6, 159.6, 3.8,
                    17.8, 4.3, 7.8, -11.6, -0.4, 17.8, 145.6, 8.4,
                    14.0, 3.6, 163.2, -4.0, 167.2, 4.7, 162.5)
T03A_DPL_REL_PRINTED = (20, 19, -7, -2, 18, 11,
                        5, 37, 10, -21, -1, 3, 45, 23,
                        36, 5, 38, -3, 54, 9, 63)

# The one row where the reference is not merely rounded but wrong. Group result
# is 422.3 against a prior year of 236.1, so ΔPY is +186.2 and +79%; the
# reference prints +162.5 and +69%, which are the ΔPL figures repeated - the
# percentage was then computed from the wrong number, so the two agree with each
# other and with nothing else.
#
# It contradicts the same table twice over: the walk (ΔPY of the result after tax
# less ΔPY of the profit to other investors is 192.9 - 6.7 = 186.2) and the row's
# own three figures. Reproducing it would make our recreation disagree with its
# own arithmetic, which is the one thing this project has consistently refused -
# so the derived value is drawn and this note is the record.
T03A_SOURCE_ERROR = {"row": 20, "column": "ΔPY",
                     "printed": 162.5, "derived": 186.2}

T03A_DPY, T03A_DPY_REL = _variances(T03A_AC, T03A_PY)
T03A_DPL, T03A_DPL_REL = _variances(T03A_AC, T03A_PL)


def _without_error(printed):
    """The printed labels, with the wrong cell blanked so the derived one shows.

    The standing rule is that the SVG prints what the original printed, which is
    how a derived 37.4 is drawn as the 37.3 the reference shows. It cannot apply
    to a cell the reference got wrong: printing +162.5 there would put a number
    on the page that the same page's own three figures contradict.
    """
    out = list(printed)
    out[T03A_SOURCE_ERROR["row"]] = None
    return tuple(out)


def _statement_columns(py, pl, ac, dpy, dpy_rel, dpy_rel_printed,
                       dpl, dpl_rel, dpl_rel_printed):
    """A statement's seven columns: three scenarios and two variance pairs."""
    return (
        Tier(key="m", label="", kind="measure", number_format="{:,.1f}",
             series=(Series("PY", py, label="PY"),
                     Series("PL", pl, label="PL"),
                     Series("AC", ac, label="AC"))),
        Tier(key="dpy", label="ΔPY", kind="variance_abs", reference="PY",
             number_format="{:+,.1f}", printed=_without_error(T03A_DPY_PRINTED),
             series=(Series("AC", dpy),)),
        Tier(key="dpyp", label="", kind="variance_rel", reference="PY",
             number_format="{:+,.0f}%", printed=_without_error(dpy_rel_printed),
             series=(Series("AC", dpy_rel),)),
        Tier(key="dpl", label="ΔPL", kind="variance_abs", reference="PL",
             number_format="{:+,.1f}", printed=T03A_DPL_PRINTED,
             series=(Series("AC", dpl),)),
        Tier(key="dplp", label="", kind="variance_rel", reference="PL",
             number_format="{:+,.0f}%", printed=dpl_rel_printed,
             series=(Series("AC", dpl_rel),)),
    )


T03A = Template(
    id="T03",
    variant="A",
    kind="statement table",
    title=TitleBlock(
        entity="Pharmaceutical Inc.",
        measure="Profit after tax",
        unit="kEUR",
        period="2026",
        message="",
    ),
    rows=T03A_ROWS,
    categories=tuple(r.label for r in T03A_ROWS),
    category_scenarios=("AC",) * len(T03A_ROWS),
    ratio_of=(T03A_RESULT, T03A_REVENUE),
    tiers=_statement_columns(
        T03A_PY, T03A_PL, T03A_AC,
        T03A_DPY, T03A_DPY_REL, T03A_DPY_REL_PRINTED,
        T03A_DPL, T03A_DPL_REL, T03A_DPL_REL_PRINTED),
    source_ref="template-refs/T03_T03A.png",
    notes=(
        "Every line is transcribed; the gross margin row and all four variance "
        "columns are derived.",
        "Impact direction is per row, not per column: a cost line up is red and "
        "a tax line down is green, and both fall out of Row.sign.",
        "The reference's Group result ΔPY is wrong - see T03A_SOURCE_ERROR.",
    ),
)


def _check_t03a(check) -> None:
    """Tie-outs for T03A.

    A statement checks itself, and this one is checked twice over: the walk has
    to reproduce every subtotal in three scenarios, and every one of the 84
    printed variance labels has to come back from the derived values.

    The tolerance is the point of interest. The reference prints a tenth and
    computes on more, so a walk over its printed lines lands 0.1 out - and it is
    worth being precise about how often, because "within a tolerance" is where a
    real transcription error would hide.
    """
    for scenario, values in (("PY", T03A_PY), ("PL", T03A_PL), ("AC", T03A_AC)):
        spans = waterfall_spans(T03A, values)
        drift = []
        for i, (row, (lo, hi)) in enumerate(zip(T03A_ROWS, spans)):
            if row.kind != "subtotal":
                continue
            got = hi if row.spans == "zero" else abs(hi - lo)
            gap = abs(got - values[i])
            if gap > 0.15:
                drift.append(f"{row.label} walks to {got:.1f}, printed "
                             f"{values[i]:.1f}")
        check(f"{scenario}: every subtotal is the lines above it, to a rounding",
              not drift,
              "all 6 agree within 0.1" if not drift else "; ".join(drift))

    # How much of that 0.1 there is. Stated as a count rather than swallowed by
    # the tolerance above, so a transcription error that happened to land inside
    # it would show up here as a change in the number of drifting rows.
    rounded = 0
    for values in (T03A_PY, T03A_PL, T03A_AC):
        spans = waterfall_spans(T03A, values)
        for i, (row, (lo, hi)) in enumerate(zip(T03A_ROWS, spans)):
            if row.kind != "subtotal":
                continue
            got = hi if row.spans == "zero" else abs(hi - lo)
            if 1e-9 < abs(got - values[i]) <= 0.15:
                rounded += 1
    check("the drift is one rounding per scenario, carried down",
          rounded == 11,
          f"{rounded} of 18 subtotals sit 0.1 from the printed figure: prior "
          f"year picks it up at the result before tax and plan and actual at "
          f"the operating result, and each carries to every result line below")

    check(
        "the margin names the two rows it is the quotient of",
        T03A.ratio_of == (T03A_RESULT, T03A_REVENUE)
        and T03A_ROWS[T03A.ratio_of[0]].label == "Operating result"
        and T03A_ROWS[T03A.ratio_of[1]].label == "Sales revenue",
        "operating result over sales revenue, read by both renderers",
    )

    check(
        "the gross margin is the operating result over sales revenue",
        all(abs(excel_round(v[T03A_MARGIN], 1) - p) < 1e-9
            for v, p in ((T03A_PY, 36.7), (T03A_PL, 36.9), (T03A_AC, 45.3))),
        "36.7%, 36.9% and 45.3%, derived rather than typed",
    )

    check(
        "the ratio row stays out of the walk",
        T03A_ROWS[T03A_MARGIN].kind == "ratio"
        and abs(waterfall_spans(T03A, T03A_AC)[T03A_MARGIN + 1][0]
                - waterfall_spans(T03A, T03A_AC)[T03A_MARGIN - 1][1]) < 1e-9,
        "investment income starts where the operating result ended, so the "
        "margin did not advance the statement",
    )

    for name, derived, printed, tolerance in (
            ("dPY", T03A_DPY, T03A_DPY_PRINTED, 0.15),
            ("dPL", T03A_DPL, T03A_DPL_PRINTED, 0.15)):
        wrong = [f"{T03A_ROWS[i].label} derives {d:+.1f}, printed {p:+.1f}"
                 for i, (d, p) in enumerate(zip(derived, printed))
                 if abs(d - p) > tolerance and i != T03A_SOURCE_ERROR["row"]]
        check(f"every derived {name} matches the printed label to a rounding",
              not wrong,
              "all 21 agree within 0.1" if not wrong else "; ".join(wrong))

    for name, derived, printed in (
            ("dPY%", T03A_DPY_REL, T03A_DPY_REL_PRINTED),
            ("dPL%", T03A_DPL_REL, T03A_DPL_REL_PRINTED)):
        wrong = [f"{T03A_ROWS[i].label} derives {d:+.1f}, printed {p:+d}"
                 for i, (d, p) in enumerate(zip(derived, printed))
                 if abs(excel_round(d) - p) > 1 and i != T03A_SOURCE_ERROR["row"]]
        check(f"every derived {name} rounds to the printed label",
              not wrong,
              "all 21 agree" if not wrong else "; ".join(wrong))

    # The error is asserted, not just noted. If a later pass "corrects" the
    # transcription to match the reference, this fails and says why.
    row = T03A_SOURCE_ERROR["row"]
    derived = T03A_AC[row] - T03A_PY[row]
    check(
        "the reference's Group result dPY is wrong, and we know by how much",
        abs(derived - T03A_SOURCE_ERROR["derived"]) < 0.05
        and abs(T03A_DPY_PRINTED[row] - T03A_SOURCE_ERROR["printed"]) < 1e-9
        and abs(T03A_DPL_PRINTED[row] - T03A_SOURCE_ERROR["printed"]) < 1e-9,
        f"422.3 less 236.1 is {derived:+.1f}, and the reference prints "
        f"{T03A_SOURCE_ERROR['printed']:+.1f} in both variance columns - the "
        f"dPL figure repeated, with its percentage taken from it",
    )


# --------------------------------------------------------------------------- #
# T04A - a statement with its variances drawn, against plan
# --------------------------------------------------------------------------- #
#
# T03A's statement told T02A's way: the same twenty-one rows, plan and actual
# only, and the two variance columns drawn rather than printed. It is the fourth
# corner of the table tranche, and by now every part of it exists - the rows are
# a statement's, the columns are a table's, and the panels are the ones T02A
# built.
#
# Two things are its own.
#
# **The margin has no variance.** T03A prints one (+8.4%p); T04A draws nothing
# on that row at all. A drawn variance would have to be a bar in the column's
# unit, and the margin is not in that unit - so rather than mix denominations
# inside one panel the reference simply leaves it out. The series carry None
# there, and both engines already skip a None.
#
# **The figures are printed to a tenth over data with more precision**, further
# than T03A: the relative variances come back within about a quarter of a
# percentage point rather than a tenth. Maintenance is the worst - 5.6 over 22.0
# is 25.45%, printed 25.3 - which says the underlying numbers are nearer 5.58
# over 22.04. Reproducible to a rounding, not exactly, and the tie-out says so.

T04A_ROWS = (
    Row("Licences", +1, prefix="+"),
    Row("Consulting", +1, prefix="+"),
    Row("Maintenance", +1, prefix="+"),
    Row("Other revenue", +1, prefix="+"),
    Row("Sales revenue", +1, kind="subtotal", prefix="=", spans="zero"),
    Row("Other operating income", +1, prefix="+"),
    Row("Purchases", -1, indent=1),
    Row("Material expenses", -1, indent=1),
    Row("Personnel expenses", -1, indent=1),
    Row("Amortization", -1, indent=1),
    Row("Other operating expenses", -1, indent=1),
    Row("Operating expenses", -1, kind="subtotal", prefix="-", spans=5),
    Row("Operating result", +1, kind="subtotal", prefix="=", spans="zero"),
    Row("Gross margin", +1, kind="ratio"),
    Row("Investment income", +1, prefix="+"),
    Row("Financial income, net", +1, prefix="+"),
    Row("Result before tax", +1, kind="subtotal", prefix="=", spans="zero"),
    Row("Income tax", -1, prefix="-"),
    Row("Result after tax", +1, kind="subtotal", prefix="=", spans="zero"),
    Row("Profit to other investors", -1, prefix="-"),
    Row("Group result", +1, kind="subtotal", prefix="=", spans="zero"),
)

T04A_MARGIN, T04A_RESULT, T04A_REVENUE = 13, 12, 4

T04A_PL_LINES = (749.0, 75.6, 22.0, 62.8, 909.4, 40.5,
                 589.5, 55.9, 79.8, 56.0, 76.0, 857.2, 92.7, None,
                 44.7, 66.0, 203.4, 127.8, 75.6, 52.0, 23.6)
T04A_AC_LINES = (745.9, 80.1, 27.6, 65.0, 918.6, 49.7,
                 600.6, 54.0, 85.6, 47.5, 79.7, 867.3, 100.9, None,
                 53.0, 66.0, 220.0, 135.3, 84.7, 55.0, 29.7)


def _with_t04a_margin(values):
    out = list(values)
    out[T04A_MARGIN] = out[T04A_RESULT] / out[T04A_REVENUE] * 100
    return tuple(out)


T04A_PL = _with_t04a_margin(T04A_PL_LINES)
T04A_AC = _with_t04a_margin(T04A_AC_LINES)

T04A_DPL_PRINTED = (-3.1, 4.5, 5.6, 2.3, 9.2, 9.1,
                    11.1, -1.9, 5.8, -8.6, 3.7, 10.2, 8.2, None,
                    8.3, 0.0, 16.6, 7.5, 9.1, 3.0, 6.0)
T04A_DPL_REL_PRINTED = (-0.4, 5.9, 25.3, 3.6, 1.0, 22.5,
                        1.9, -3.3, 7.3, -15.3, 4.8, 1.2, 8.9, None,
                        18.7, 0.0, 8.2, 5.9, 12.0, 5.8, 25.6)


def _blank_ratio(values, index):
    """The ratio row carries no variance here, so nothing is drawn on it."""
    out = list(values)
    out[index] = None
    return tuple(out)


_dpl, _dpl_rel = _variances(T04A_AC, T04A_PL)
T04A_DPL = _blank_ratio(_dpl, T04A_MARGIN)
T04A_DPL_REL = _blank_ratio(_dpl_rel, T04A_MARGIN)


T04A = Template(
    id="T04",
    variant="A",
    kind="statement table with integrated bars",
    title=TitleBlock(
        entity="Pharmaceutical Inc.",
        measure="Profit after tax",
        unit="kEUR",
        period="2026",
        message="",
    ),
    rows=T04A_ROWS,
    categories=tuple(r.label for r in T04A_ROWS),
    category_scenarios=("AC",) * len(T04A_ROWS),
    ratio_of=(T04A_RESULT, T04A_REVENUE),
    tiers=(
        Tier(key="m", label="", kind="measure", number_format="{:,.1f}",
             series=(Series("PL", T04A_PL, label="PL"),
                     Series("AC", T04A_AC, label="AC"))),
        Tier(key="dpl", label="ΔPL", kind="variance_abs", reference="PL",
             number_format="{:+,.1f}", printed=T04A_DPL_PRINTED,
             series=(Series("AC", T04A_DPL),)),
        Tier(key="dplp", label="ΔPL%", kind="variance_rel", reference="PL",
             number_format="{:+,.1f}", printed=T04A_DPL_REL_PRINTED,
             series=(Series("AC", T04A_DPL_REL),)),
    ),
    panel_tiers=("dpl", "dplp"),
    source_ref="template-refs/T04_T04A.png",
    notes=(
        "Plan and actual only; both variance columns are drawn.",
        "The gross margin row carries no variance - a drawn bar would be in the "
        "column's unit and a margin is not.",
        "Impact is per row: a cost line up is red and a cost line down is "
        "green, which is Row.sign and nothing else.",
    ),
)


def _check_t04a(check) -> None:
    """Tie-outs for T04A.

    The same statement checks as T03A over one fewer scenario, plus the two that
    are particular to a drawn variance: that the margin really carries nothing,
    and that the impact colouring falls out of the statement's own structure
    rather than a hand-written list of which rows are costs.
    """
    for scenario, values in (("PL", T04A_PL), ("AC", T04A_AC)):
        spans = waterfall_spans(T04A, values)
        drift = []
        for i, row in enumerate(T04A_ROWS):
            if row.kind != "subtotal":
                continue
            lo, hi = spans[i]
            got = hi if row.spans == "zero" else abs(hi - lo)
            if abs(got - values[i]) > 0.15:
                drift.append(f"{row.label} walks to {got:.1f}, printed {values[i]:.1f}")
        check(f"{scenario}: every subtotal is the lines above it, to a rounding",
              not drift,
              "all 6 agree within 0.1" if not drift else "; ".join(drift))

    check(
        "the gross margin is the operating result over sales revenue",
        abs(excel_round(T04A_PL[T04A_MARGIN], 1) - 10.2) < 1e-9
        and abs(excel_round(T04A_AC[T04A_MARGIN], 1) - 11.0) < 1e-9,
        "10.2% against plan and 11.0% actual, derived rather than typed",
    )

    check(
        "the margin carries no variance, in either column",
        T04A_DPL[T04A_MARGIN] is None and T04A_DPL_REL[T04A_MARGIN] is None,
        "nothing is drawn on that row, as the reference draws nothing",
    )

    wrong = [f"{T04A_ROWS[i].label} derives {d:+.1f}, printed {p:+.1f}"
             for i, (d, p) in enumerate(zip(T04A_DPL, T04A_DPL_PRINTED))
             if d is not None and abs(d - p) > 0.15]
    check("every derived dPL matches the printed label to a rounding",
          not wrong, "all 20 agree within 0.1" if not wrong else "; ".join(wrong))

    # A quarter of a point, not a tenth. Stated as the actual worst case rather
    # than as a round tolerance, so that a transcription error large enough to
    # matter cannot hide inside a number chosen for comfort.
    worst, where = 0.0, ""
    for i, (d, p) in enumerate(zip(T04A_DPL_REL, T04A_DPL_REL_PRINTED)):
        if d is None:
            continue
        if abs(d - p) > worst:
            worst, where = abs(d - p), T04A_ROWS[i].label
    check("every derived dPL% matches the printed label to a quarter of a point",
          worst <= 0.3,
          f"worst is {where} at {worst:.2f} points - the reference prints a "
          f"tenth over data with more precision, so 5.6 over 22.0 reads 25.3 "
          f"where 25.45 is what the printed figures give")

    adverse = [i for i, r in enumerate(T04A_ROWS)
               if not r.higher_is_better and T04A_DPL[i] is not None
               and T04A_DPL[i] > 0]
    check(
        "cost increases classify as adverse without being listed as such",
        adverse == [6, 8, 10, 11, 17, 19],
        "purchases, personnel, other operating expenses and their total, plus "
        "income tax and the profit paid away - six rows, every one an increase "
        "on a subtracting line",
    )


# --------------------------------------------------------------------------- #
# C01A - stacked columns, Alpha Software Corporation net sales 2022..2026
# --------------------------------------------------------------------------- #
#
# The first *structure* chart in the library. Everything before it stacked
# scenarios or waterfall steps; here the bands are business areas, and every one
# of them is an actual - which is why they are Segments rather than Series, and
# why the scenario notation has to be carried by something else. It is carried
# by the column: 2022 to 2025 are actuals and 2026 is plan.
#
# Three panels share one scale. That is the whole design: 101.9 is drawn the
# same height in all three, so a reader can carry the 2025 total from the time
# series into either breakdown of it. Two of the tie-outs below are exactly that
# claim.
#
# Transcription note: the reference suppresses a label where the band is too
# thin to hold one, which is a required feature of the template rather than an
# omission - so Training in 2022 and 2023 has no printed figure and is recovered
# from the column total, 74.9 - 74.0 = 0.9 and 79.9 - 77.0 = 2.9. Both are then
# checked back against the printed totals.

C01A_AREAS = ("Software", "Service", "Training", "Consulting", "Other")
C01A_YEARS = ("2022", "2023", "2024", "2025", "2026")

# Printed everywhere it fits. Training's first two are the two the reference
# leaves off, derived from the totals it does print.
C01A_BY_AREA = {
    "Software":   (29.1, 31.9, 40.8, 40.2, 49.3),
    "Service":    (16.0, 18.8, 24.2, 26.7, 28.7),
    "Training":   (0.9, 2.9, 5.1, 7.7, 8.6),
    "Consulting": (22.4, 22.4, 25.3, 22.4, 20.2),
    "Other":      (6.5, 3.9, 5.2, 4.9, 7.3),
}
C01A_AREA_TOTALS = (74.9, 79.9, 100.6, 101.9, 114.1)

C01A_BY_INDUSTRY = (
    ("Public sector", 35.3), ("Finance industry", 19.8), ("Consumer goods", 14.3),
    ("Retail", 15.9), ("Engineering", 7.4), ("Other", 9.2),
)
C01A_BY_REGION = (
    ("USA", 32.5), ("Germany", 17.7), ("Japan", 12.4),
    ("Rest of Europe", 10.7), ("Rest of World", 28.6),
)

C01A = Template(
    id="C01",
    variant="A",
    kind="stacked column chart",
    title=TitleBlock(
        entity="Alpha Software Corporation",
        measure="Net sales",
        unit="kEUR",
        period="2022..2026",
        message=("Planned net sales in 2026 will increase by 12.2 kEUR (+12%) "
                 "mainly due to software growth of 9.1 kEUR (+23%)"),
    ),
    categories=C01A_YEARS,
    category_scenarios=("AC", "AC", "AC", "AC", "PL"),
    tiers=(),
    structure_panels=(
        StructurePanel(
            key="area", label="By business area",
            categories=C01A_YEARS,
            category_scenarios=("AC", "AC", "AC", "AC", "PL"),
            segments=tuple(Segment(name, C01A_BY_AREA[name]) for name in C01A_AREAS),
            legend_side="left", legend_at=0,
            printed_totals=C01A_AREA_TOTALS),
        StructurePanel(
            key="industry", label="By industry",
            categories=("2025",), category_scenarios=("AC",),
            segments=tuple(Segment(n, (v,)) for n, v in C01A_BY_INDUSTRY),
            legend_side="left"),
        StructurePanel(
            key="region", label="By region",
            categories=("2025",), category_scenarios=("AC",),
            segments=tuple(Segment(n, (v,)) for n, v in C01A_BY_REGION),
            legend_side="right"),
    ),
    annotations=(
        # The two brackets the message names, each spanning one year's growth.
        Annotation("bracket", ("area", 4), text="+12.2"),
        Annotation("bracket", ("area:Software", 4), text="+9.1"),
        # The two comment markers, each on the band its note is about.
        Annotation("oval", ("area:Consulting", 4), ref=1),
        Annotation("oval", ("region:Rest of Europe", 0), ref=2),
    ),
    comments=(
        Comment(1, "2026 Plan Consulting",
                "Increased net sales of 12.2 kEUR compared to 2025 because of "
                "Beta Corp. project completed on beginning in February 2025."),
        Comment(2, "2025 Rest of Europe",
                "The largest shares are 5.1 kEUR from UK and 3.4 kEUR from "
                "the Netherlands"),
    ),
    source_ref="template-refs/C01_01A.png",
    notes=(
        "Bands are structure categories, not scenarios; the scenario is carried "
        "by the column - four actuals and a plan.",
        "All three panels share one scale, which is what makes 101.9 in the "
        "time series and 101.9 in each breakdown the same height.",
        "Training 2022 and 2023 are derived: the reference suppresses a label "
        "where the band is too thin to hold one.",
    ),
)


def _check_c01a(check) -> None:
    """Tie-outs for C01A.

    The strongest is the last one. Two of the three panels are breakdowns of a
    number the third one draws, so if all three add to 101.9 then the two
    transcribed breakdowns and the 2025 column of the time series are all right
    together - and that is the comparison the whole layout exists to make.
    """
    area = C01A.structure_panels[0]
    wrong = [f"{C01A_YEARS[i]} sums to {area.total(i):.1f}, printed {p:.1f}"
             for i, p in enumerate(C01A_AREA_TOTALS)
             if abs(area.total(i) - p) > 0.05]
    check("every year's bands sum to the total printed over the column",
          not wrong, "all 5 agree exactly" if not wrong else "; ".join(wrong))

    # The two suppressed labels, recovered and stated. If either is wrong its
    # column total fails above, which is what makes deriving them safe.
    check(
        "the two suppressed Training bands are recovered from the totals",
        abs(C01A_BY_AREA["Training"][0] - 0.9) < 1e-9
        and abs(C01A_BY_AREA["Training"][1] - 2.9) < 1e-9,
        "0.9 in 2022 and 2.9 in 2023, too thin for the reference to label",
    )

    industry = C01A.structure_panels[1]
    region = C01A.structure_panels[2]
    totals = (area.total(3), industry.total(0), region.total(0))
    check(
        "all three panels are breakdowns of the same 101.9",
        all(abs(t - 101.9) < 0.05 for t in totals),
        f"{totals[0]:.1f} by business area, {totals[1]:.1f} by industry and "
        f"{totals[2]:.1f} by region - the comparison the layout is making",
    )

    # The message states both bracket figures; they have to come from the data
    # rather than from the sentence.
    growth = area.total(4) - area.total(3)
    software = (C01A_BY_AREA["Software"][4] - C01A_BY_AREA["Software"][3])
    check(
        "both bracketed figures are what the columns actually say",
        abs(growth - 12.2) < 0.05 and abs(software - 9.1) < 0.05,
        f"the plan is {growth:+.1f} kEUR over 2025 ({growth / area.total(3) * 100:+.0f}%) "
        f"and software {software:+.1f} ({software / C01A_BY_AREA['Software'][3] * 100:+.0f}%), "
        f"which is the message word for word",
    )

    check(
        "2026 is the only planned column",
        [i for i, s in enumerate(C01A.category_scenarios) if s == "PL"] == [4],
        "four actuals and a plan, which is what the column fill has to say",
    )


# --------------------------------------------------------------------------- #
# C02A - stacked bars, Pharmaceutical Inc. net sales by country 2025
# --------------------------------------------------------------------------- #
#
# C01A turned on its side, and then given a hierarchy. The bands are the same
# kind of thing - structure categories, three sales channels - but the rows are
# countries with two subtotals among them, so the template needs both of the
# things this project has built: Segments for the bands and Rows for the
# hierarchy.
#
# **The subtotal rows are not bars.** Europe and World are drawn as a row of
# swatch-and-figure pairs, one per channel, then the total and its share. That
# is the integrated legend this template requires - the totals row doubles as
# the key, so the reader never has to match a grey to a caption, and it is why
# there is no legend box anywhere on the page.
#
# Five bands are too thin for the reference to label, and all five are recovered
# from the subtotals rather than guessed: Ireland and Denmark's wholesale at 7,
# Norway's direct and retail at 4 and 5, and Canada's retail at 5. Each falls
# out of a column that has to add to a printed figure, and _check_c02a proves
# all three columns land exactly.

C02A_CHANNELS = ("Direct", "Retail", "Wholesale")

C02A_ROWS = (
    Row("Switzerland", +1), Row("Germany", +1), Row("Austria", +1),
    Row("France", +1), Row("UK", +1), Row("Poland", +1), Row("Spain", +1),
    Row("Portugal", +1), Row("Sweden", +1), Row("Ireland", +1),
    Row("Finland", +1), Row("Denmark", +1), Row("Norway", +1),
    Row("Rest of Europe", +1),
    Row("Europe", +1, kind="subtotal", spans=14),
    Row("USA", +1), Row("Brazil", +1), Row("Japan", +1), Row("Canada", +1),
    Row("Rest of world", +1),
    Row("World", +1, kind="subtotal", spans="zero"),
)

# Element rows only, in row order; the two subtotals are derived below. None
# marks a band the reference is too small to label - recovered, then checked.
C02A_ELEMENTS = {
    "Direct": (172, 160, 102, 178, 79, 87, 83, 61, 42, 56, 48, 7, 4, 41,
               82, 52, 22, 15, 60),
    "Retail": (136, 135, 188, 88, 105, 100, 83, 44, 42, 34, 35, 24, 5, 88,
               96, 36, 16, 5, 48),
    "Wholesale": (188, 142, 125, 78, 84, 66, 58, 12, 17, 7, 13, 21, 13, 47,
                  109, 27, 26, 12, 29),
}
# The bands the reference leaves unlabelled, by (channel, row label).
C02A_SUPPRESSED = (("Wholesale", "Ireland"), ("Direct", "Denmark"),
                   ("Direct", "Norway"), ("Retail", "Norway"),
                   ("Retail", "Canada"))

# What the reference prints on the two subtotal rows, and the share it gives
# each. Targets, not inputs.
C02A_PRINTED_SUBTOTALS = {
    "Europe": {"Direct": 1120, "Retail": 1107, "Wholesale": 871, "total": 3098},
    "World": {"Direct": 1351, "Retail": 1308, "Wholesale": 1074, "total": 3733},
}
C02A_PRINTED_SHARES = {"Europe": 83, "World": 100}


def _c02a_column(channel: str) -> tuple[float, ...]:
    """One channel across every row, with the two subtotals walked in.

    The subtotals are derived rather than typed, which is what makes the five
    recovered bands safe: each one is the difference between a column of
    elements and a total the reference prints, so getting one wrong breaks a
    figure we did not supply.
    """
    values, out, running, taken = C02A_ELEMENTS[channel], [], 0.0, 0
    for row in C02A_ROWS:
        if row.kind == "subtotal":
            if row.spans == "zero":
                out.append(running)
            else:
                out.append(sum(values[taken - int(row.spans):taken]))
        else:
            out.append(values[taken])
            running += values[taken]
            taken += 1
    return tuple(out)


C02A_BY_CHANNEL = {c: _c02a_column(c) for c in C02A_CHANNELS}

C02A = Template(
    id="C02",
    variant="A",
    kind="stacked bar chart",
    orientation="horizontal",
    title=TitleBlock(
        entity="Pharmaceutical Inc.",
        measure="Net sales",
        unit="kCHF",
        period="2025",
        message=("In Europe we achieved 3 098 kCHF (83%) of worldwide net "
                 "sales (3 733 kCHF), USA net sales of 287 kCHF presents the "
                 "biggest share outside of Europe (8%)"),
    ),
    rows=C02A_ROWS,
    categories=tuple(r.label for r in C02A_ROWS),
    category_scenarios=("AC",) * len(C02A_ROWS),
    tiers=(),
    structure_panels=(
        StructurePanel(
            key="channel", label="",
            categories=tuple(r.label for r in C02A_ROWS),
            category_scenarios=("AC",) * len(C02A_ROWS),
            segments=tuple(Segment(c, C02A_BY_CHANNEL[c]) for c in C02A_CHANNELS),
            legend_side="left"),
    ),
    annotations=(
        Annotation("arrow", ("channel", 0)),      # the largest country
        Annotation("arrow", ("channel", 15)),     # the largest outside Europe
        Annotation("oval", ("channel:Retail", 2), ref=1),
        Annotation("oval", ("channel:Wholesale", 15), ref=2),
        Annotation("oval", ("channel:Wholesale", 18), ref=3),
    ),
    comments=(
        Comment(1, "Austria Retail:",
                'Vienna "Golden Line" is 98 kCHF of 188 kCHF'),
        Comment(2, "USA wholesale:",
                "50% above plan because of new channel partner at the East Coast"),
        Comment(3, "Canada:",
                "40% below plan because of the delayed Alpha project"),
    ),
    source_ref="template-refs/C02_02A.png",
    notes=(
        "Rows are sorted by total net sales, which the subject line states - "
        "the ordering is part of what the chart says.",
        "The two subtotal rows are drawn as swatch-and-figure pairs rather "
        "than bars: the totals row is this template's integrated legend.",
        "Five bands are unlabelled in the reference and recovered from the "
        "subtotals; see check_ties.",
    ),
)


def _check_c02a(check) -> None:
    """Tie-outs for C02A.

    Three columns, each of which has to add to a figure the reference prints
    twice - once for Europe and once for the World. That is what licenses the
    five recovered bands: they are not guesses to be taken on trust but the only
    values that make three independent sums come out right.
    """
    panel = C02A.structure_panels[0]
    labels = [r.label for r in C02A_ROWS]

    for name, printed in C02A_PRINTED_SUBTOTALS.items():
        row = labels.index(name)
        wrong = [f"{c} walks to {C02A_BY_CHANNEL[c][row]:g}, printed {printed[c]:g}"
                 for c in C02A_CHANNELS
                 if abs(C02A_BY_CHANNEL[c][row] - printed[c]) > 1e-9]
        check(f"{name}: every channel sums to the figure printed on the row",
              not wrong,
              "all 3 agree exactly" if not wrong else "; ".join(wrong))
        total = panel.total(row)
        check(f"{name}: the three channels sum to the printed total",
              abs(total - printed["total"]) < 1e-9,
              f"{total:,.0f} kCHF, as printed")

    check(
        "the five suppressed bands are the only values that make those sums work",
        len(C02A_SUPPRESSED) == 5,
        "Ireland and Denmark wholesale at 7, Norway direct and retail at 4 and "
        "5, Canada retail at 5 - each recovered from a column the reference "
        "totals but does not itemise",
    )

    europe = panel.total(labels.index("Europe"))
    world = panel.total(labels.index("World"))
    share = europe / world * 100
    check("the shares the reference circles are what the totals give",
          abs(excel_round(share) - C02A_PRINTED_SHARES["Europe"]) < 1e-9,
          f"{europe:,.0f} of {world:,.0f} is {share:.1f}%, printed 83%")

    usa = panel.total(labels.index("USA"))
    check("the USA share the message names is what its bar says",
          abs(excel_round(usa / world * 100) - 8) < 1e-9,
          f"{usa:,.0f} of {world:,.0f} is {usa / world * 100:.1f}%, printed 8%")

    # Split on the subtotals rather than at a counted position: the blocks are
    # what the hierarchy says they are, and counting them by hand is how this
    # check first failed on thirteen European rows it had been told were twelve.
    blocks, current = [], []
    for i, row in enumerate(C02A_ROWS):
        if row.kind == "subtotal":
            blocks.append(current)
            current = []
        elif row.label not in ("Rest of Europe", "Rest of world"):
            current.append(panel.total(i))
    unsorted = [n for n, b in enumerate(blocks)
                if b != sorted(b, reverse=True)]
    check(
        "rows are sorted by total, which the subject line claims",
        not unsorted,
        f"descending within both blocks - {len(blocks[0])} European rows and "
        f"{len(blocks[1])} outside - with the two residual rows placed last by "
        f"meaning rather than by size"
        if not unsorted else f"block(s) {unsorted} are out of order",
    )


# --------------------------------------------------------------------------- #
# C07C - line chart, Alpha Software Corp. net sales 2025
# --------------------------------------------------------------------------- #
#
# The first template drawn as lines, and the first with four series on one pair
# of axes: a moving annual total across the top, and three cumulative lines -
# plan, actual to August, forecast from September - climbing to meet it. Under
# them sits a tier of monthly columns, on the same scale, which is what the
# lines are the running sums of.
#
# Transcription was the hardest in the project so far and worth recording why.
# The reference labels only where a label fits, so the printed figures are a
# subset scattered across four lines, and two of them sit close enough to be
# read against the wrong series - 1 902 and 1 921 at November are the plan and
# the forecast, and getting them the wrong way round is not visible on the page.
#
# What settled it was measurement rather than reading: the plan columns at
# November and December measure 107 and 86 units tall, which only works if the
# plan cumulative runs ... 1 794, 1 902, 1 990. The other assignment would need
# them at 127 and 69. Everything else follows from three sources that have to
# agree - the monthly columns, the cumulative lines they sum to, and the two
# bracketed variances.
#
# The structural tie-out is the last line of the chart: a moving annual total in
# December *is* the year, so MAT must land exactly on the forecast cumulative.

C07C_MONTHS = ("Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
C07C_SCENARIOS = ("AC",) * 8 + ("FC",) * 4

# The monthly tier, transcribed from the column labels and confirmed against the
# column heights. Actual through August, forecast after.
C07C_MONTHLY = (145, 154, 178, 189, 155, 165, 176, 168, 211, 188, 189, 155)
C07C_MONTHLY_PL = (145, 240, 217, 217, 171, 169, 176, 189, 143, 127, 108, 88)

# What the reference prints on the cumulative lines. None where it prints
# nothing - which is most of the actual line, because the labels would collide.
C07C_CUM_PRINTED = (None, None, 478, 667, 823, 988, 1164, 1332,
                    1544, 1732, 1921, 2077)
C07C_CUM_PL_PRINTED = (145, 385, 602, 819, 990, 1159, 1335, 1524,
                       1667, 1794, 1902, 1990)

# The moving annual total. Six of the twelve are printed; the rest are recovered
# from the marker positions, which the scale makes good to about three units.
# MAT is the one series here that cannot be derived from anything else on the
# page - it reaches back into 2024 - so it is transcribed and then checked at
# the one point where it must agree with the rest, December.
C07C_MAT = (1986, 1990, 2031, 2072, 2072, 2069, 2087, 2084, 2111, 2118, 2118, 2077)
C07C_MAT_PRINTED = (1986, None, 2031, 2072, None, None, 2087, None,
                    2111, None, None, 2077)


def _running(values):
    out, total = [], 0.0
    for v in values:
        total += v
        out.append(total)
    return tuple(out)


C07C_CUM = _running(C07C_MONTHLY)
C07C_CUM_PL = _running(C07C_MONTHLY_PL)

C07C = Template(
    id="C07",
    variant="C",
    kind="line chart",
    title=TitleBlock(
        entity="Alpha Software Corp.",
        measure="Net sales",
        unit="kUSD",
        period="2025",
        message=("Till the end of August our net sales was 192 kUSD lower than "
                 "planned, because … However, we estimate that full year "
                 "net sales will be higher than budgeted"),
    ),
    categories=C07C_MONTHS,
    category_scenarios=C07C_SCENARIOS,
    tiers=(
        Tier(key="mat", label="MAT", kind="measure", number_format="{:,.0f}",
             printed=C07C_MAT_PRINTED,
             series=(Series("AC", _split(C07C_MAT, "AC", C07C_SCENARIOS), label="MAT"),
                     Series("FC", _split(C07C_MAT, "FC", C07C_SCENARIOS), label="MAT"))),
        # One tier per line, not one per group: each carries its own printed
        # labels, and the reference labels the three cumulatives differently -
        # the plan at every month, the actual only from March, the forecast
        # throughout. A shared list would print each line's figures on all
        # three.
        Tier(key="cum_pl", label="PL", kind="measure", number_format="{:,.0f}",
             printed=C07C_CUM_PL_PRINTED,
             series=(Series("PL", C07C_CUM_PL, label="PL"),)),
        Tier(key="cum_ac", label="AC", kind="measure", number_format="{:,.0f}",
             printed=_split(C07C_CUM_PRINTED, "AC", C07C_SCENARIOS),
             series=(Series("AC", _split(C07C_CUM, "AC", C07C_SCENARIOS), label="AC"),)),
        Tier(key="cum_fc", label="FC", kind="measure", number_format="{:,.0f}",
             printed=_split(C07C_CUM_PRINTED, "FC", C07C_SCENARIOS),
             series=(Series("FC", _split(C07C_CUM, "FC", C07C_SCENARIOS), label="FC"),)),
        Tier(key="monthly", label="", kind="measure", number_format="{:,.0f}",
             series=(Series("PL", C07C_MONTHLY_PL, label="PL"),
                     Series("AC", _split(C07C_MONTHLY, "AC", C07C_SCENARIOS), label="AC"),
                     Series("FC", _split(C07C_MONTHLY, "FC", C07C_SCENARIOS), label="FC"))),
    ),
    annotations=(
        Annotation("bracket", ("cum", 7), text="-192"),
        Annotation("bracket", ("cum", 11), text="+87"),
    ),
    source_ref="template-refs/C07_07C.png",
    notes=(
        "Four series on one scale: a moving annual total, and the plan, actual "
        "and forecast cumulatives the monthly tier sums to.",
        "MAT is the only series that cannot be derived from the page - it "
        "reaches into 2024 - so it is transcribed and checked at December.",
        "The reference labels only where a label fits; the unprinted values are "
        "recovered from marker positions and checked against the columns.",
    ),
)


def _check_c07c(check) -> None:
    """Tie-outs for C07C.

    Three sources have to agree: the monthly columns, the cumulative lines they
    are the running sums of, and the two variances the reference brackets. Any
    one of them alone could be mis-transcribed without showing; together they
    pin the whole chart.
    """
    wrong = [f"{C07C_MONTHS[i]} sums to {c:,.0f}, printed {p:,.0f}"
             for i, (c, p) in enumerate(zip(C07C_CUM, C07C_CUM_PRINTED))
             if p is not None and abs(c - p) > 4]
    check("the cumulative line is the running sum of the monthly columns",
          not wrong,
          "all 10 printed points agree within 4 - the reference prints a "
          "rounded month and cumulates the unrounded one"
          if not wrong else "; ".join(wrong))

    wrong = [f"{C07C_MONTHS[i]} sums to {c:,.0f}, printed {p:,.0f}"
             for i, (c, p) in enumerate(zip(C07C_CUM_PL, C07C_CUM_PL_PRINTED))
             if abs(c - p) > 1e-9]
    check("the plan cumulative is exactly the plan months",
          not wrong,
          "all 12 agree exactly - the plan months were derived from it"
          if not wrong else "; ".join(wrong))

    # The structural one. A moving annual total in December is the year, so it
    # has to land on the forecast cumulative - two series transcribed from
    # different parts of the page, meeting at one point.
    check(
        "the moving annual total lands on the year it is a total of",
        abs(C07C_MAT[-1] - C07C_CUM[-1]) <= 4,
        f"MAT December {C07C_MAT[-1]:,} against a cumulative "
        f"{C07C_CUM[-1]:,.0f} - the same year counted two ways",
    )

    august = C07C_CUM[7] - C07C_CUM_PL[7]
    check("the bracketed shortfall is what the two lines say at August",
          abs(august - (-192)) <= 3,
          f"{august:+,.0f} kUSD against a printed -192")

    full_year = C07C_CUM[-1] - C07C_CUM_PL[-1]
    check("the bracketed full-year gain is what the two lines say at December",
          abs(full_year - 87) <= 4,
          f"{full_year:+,.0f} kUSD against a printed +87")

    overlap = [C07C_MONTHS[i] for i in range(12)
               if (_split(C07C_MONTHLY, "AC", C07C_SCENARIOS)[i] is not None
                   and _split(C07C_MONTHLY, "FC", C07C_SCENARIOS)[i] is not None)]
    check("actual and forecast do not overlap",
          not overlap,
          "eight actual months and four forecast, and no month is both")

    printed = [p for p in C07C_MAT_PRINTED if p is not None]
    check("every printed MAT label is reproduced",
          all(abs(C07C_MAT[i] - p) < 1e-9
              for i, p in enumerate(C07C_MAT_PRINTED) if p is not None),
          f"{len(printed)} of 12 are labelled; the other six are recovered from "
          f"marker positions, good to about three units")


# --------------------------------------------------------------------------- #
# C08H - area chart, Alpha Software Corporation raw material 2023..2027
# --------------------------------------------------------------------------- #
#
# C07's axis with a fill, and a stock measure instead of a flow. Three tiers on
# one page: what went in and what came out each quarter, the net of the two, and
# the level that results - which is the only one of the three that is not a
# movement, and the reason the template exists.
#
# Everything here is determined by two transcribed series and one opening
# balance. Increase and decrease are read off the bottom tier; the change tier
# is their difference; the inventory line is the opening 11 tons with the
# changes accumulated onto it. Twenty-one printed levels then have to come back
# out, and they do, exactly - which is the strongest tie-out in the library
# after T01B's, and unlike T01B's it is a *recurrence*: one wrong quarter throws
# every level after it, so a single error cannot hide.
#
# The scenario runs three ways across one line: eleven quarters of actual, one
# forecast quarter, then eight of plan. The area under the forecast quarter is
# hatched and the plan markers are hollow, so the line changes what it claims
# without changing what it is.

C08H_QUARTERS = tuple(f"Q{q} {y}" for y in (2023, 2024, 2025, 2026, 2027)
                      for q in (1, 2, 3, 4))
C08H_SCENARIOS = ("AC",) * 11 + ("FC",) + ("PL",) * 8

C08H_INCREASE = (3, 2, 5, 3, 6, 3, 5, 4, 2, 1, 0, 5,
                 3, 4, 4, 5, 4, 4, 4, 6)
C08H_DECREASE = (0, 3, 3, 2, 4, 2, 3, 3, 4, 4, 6, 3,
                 2, 2, 2, 6, 2, 7, 2, 5)

# The stock at the start of the first quarter. Everything else follows.
C08H_OPENING = 11

# The twenty-one levels the reference prints, opening included. Targets, not
# inputs - the recurrence has to reproduce every one.
C08H_LEVELS_PRINTED = (11, 14, 13, 15, 16, 18, 19, 21, 22, 20, 17, 11,
                       13, 14, 16, 18, 17, 19, 16, 18, 19)

C08H_CHANGE = tuple(i - d for i, d in zip(C08H_INCREASE, C08H_DECREASE))


def _accumulate(opening, changes):
    """The level at the end of each quarter, from an opening balance."""
    out, level = [], opening
    for c in changes:
        level += c
        out.append(level)
    return tuple(out)


C08H_LEVELS = _accumulate(C08H_OPENING, C08H_CHANGE)   # 20 closing levels

C08H = Template(
    id="C08",
    variant="H",
    kind="area chart",
    title=TitleBlock(
        entity="Alpha Software Corporation",
        measure="Raw material",
        unit="tons",
        period="2023..2027",
        message=("Until end of 2025 we plan a slight increase "
                 "of our raw material stock to 13 tons"),
    ),
    categories=C08H_QUARTERS,
    category_scenarios=C08H_SCENARIOS,
    tiers=(
        Tier(key="change", label="Inventory change", kind="measure",
             number_format="{:+,.0f}",
             series=(Series("AC", _split(C08H_CHANGE, "AC", C08H_SCENARIOS)),
                     Series("FC", _split(C08H_CHANGE, "FC", C08H_SCENARIOS)),
                     Series("PL", _split(C08H_CHANGE, "PL", C08H_SCENARIOS)))),
        Tier(key="level", label="Inventory", kind="measure",
             number_format="{:,.0f}",
             series=(Series("AC", _split(C08H_LEVELS, "AC", C08H_SCENARIOS)),
                     Series("FC", _split(C08H_LEVELS, "FC", C08H_SCENARIOS)),
                     Series("PL", _split(C08H_LEVELS, "PL", C08H_SCENARIOS)))),
        Tier(key="flow", label="Increase", kind="measure",
             number_format="{:,.0f}",
             series=(Series("AC", _split(C08H_INCREASE, "AC", C08H_SCENARIOS)),
                     Series("FC", _split(C08H_INCREASE, "FC", C08H_SCENARIOS)),
                     Series("PL", _split(C08H_INCREASE, "PL", C08H_SCENARIOS)))),
        Tier(key="outflow", label="Decrease", kind="measure",
             number_format="{:,.0f}",
             series=(Series("AC", _split(C08H_DECREASE, "AC", C08H_SCENARIOS)),
                     Series("FC", _split(C08H_DECREASE, "FC", C08H_SCENARIOS)),
                     Series("PL", _split(C08H_DECREASE, "PL", C08H_SCENARIOS)))),
    ),
    annotations=(
        # The level the message names, at the end of 2025.
        Annotation("oval", ("level", 11)),
    ),
    source_ref="template-refs/C08_08H.png",
    notes=(
        "Two transcribed series and an opening balance determine the whole "
        "chart; the twenty-one printed levels are the check.",
        "One line, three scenarios: eleven actual quarters, one forecast, then "
        "eight of plan.",
        "The level is a stock, so it is drawn at quarter boundaries - twenty "
        "closings plus the opening - while the movements are drawn in the "
        "quarters they belong to.",
    ),
)


def _check_c08h(check) -> None:
    """Tie-outs for C08H.

    A recurrence rather than a set of independent sums, which makes it strict:
    the level at the end of 2027 depends on every quarter before it, so a single
    mistyped movement moves every level after it and cannot be absorbed.
    """
    wrong = [f"{C08H_QUARTERS[i]}: {inc} less {dec} is {inc - dec:+d}"
             for i, (inc, dec) in enumerate(zip(C08H_INCREASE, C08H_DECREASE))
             if inc - dec != C08H_CHANGE[i]]
    check("the change tier is the increase less the decrease",
          not wrong, "all 20 agree exactly" if not wrong else "; ".join(wrong))

    levels = (C08H_OPENING,) + C08H_LEVELS
    wrong = [f"{i}: walks to {a}, printed {b}"
             for i, (a, b) in enumerate(zip(levels, C08H_LEVELS_PRINTED))
             if a != b]
    check(
        "every printed level comes back from the opening balance and the flows",
        not wrong,
        f"all {len(levels)} agree exactly - 11 tons at the start, then twenty "
        f"quarters of movement" if not wrong else "; ".join(wrong),
    )

    # The one the message names. It is a *boundary* value - the level at the end
    # of 2025 - which is why the reference circles it on the year rule.
    end_2025 = levels[12]
    check("the level the message names is the one the chart circles",
          end_2025 == 13,
          f"{end_2025} tons at the end of 2025, which is what the sentence says")

    counts = {s: C08H_SCENARIOS.count(s) for s in ("AC", "FC", "PL")}
    check(
        "one line carries three scenarios, in the order a plan is made",
        counts == {"AC": 11, "FC": 1, "PL": 8}
        and C08H_SCENARIOS[10:13] == ("AC", "FC", "PL"),
        "eleven actual quarters, then the forecast quarter that closes the "
        "year, then two planned years",
    )

    check(
        "no quarter takes stock out that it never had",
        all(level >= 0 for level in levels),
        f"the level runs {min(levels)} to {max(levels)} tons and never goes "
        f"negative",
    )


# --------------------------------------------------------------------------- #
# C10D - bubble chart, Alpha Corp. strategic business units 2025
# --------------------------------------------------------------------------- #
#
# The first template in the library with no category axis at all. Every other
# chart here puts a name on one axis and a number on the other; this one puts a
# number on both, and a third number in the size of the mark. Nine business
# units are drawn twice - where they stood last year and where they stand now -
# and two acquisitions are drawn once, because a unit bought in December has no
# prior year inside this group to be compared with.
#
# Nothing here is derived. A portfolio chart states three independent measures
# per point and no arithmetic connects them: market attractiveness is an index
# somebody scored, relative market share is this unit's sales over the largest
# competitor's, and net sales is money. So the tie-outs cannot check sums the
# way a waterfall's can. What they check instead is that the picture supports
# the sentence written over it, and that the transcription is self-consistent
# against the one measurable relationship the chart does contain - the area law.
#
# Every coordinate was recovered by fitting a circle to each bubble's outer arc
# in the reference render - the arc that borders paper, so it genuinely lies on
# the disc rather than on an overlap. Centre and radius were fitted together,
# nothing assumed. All forty coordinates then landed within 0.005 of a
# two-decimal value, which is what says the transcription is right rather than
# merely close.

C10D_AXES = (
    Axis("Market attractiveness", 0.0, 1.00, 0.50),
    Axis("Relative market share", 0.0, 1.25, 0.25),
)

# name, then (attractiveness, relative share, net sales in mEUR) for PY and AC.
# Ordered by prior-year attractiveness, which is only a reading order - an XY
# chart's rows have no position on the page.
C10D_SBUS = (
    ("RS1", (0.11, 1.10,  4.0), (0.16, 1.00,  9.0)),
    ("VBN", (0.13, 0.68,  2.0), (0.10, 0.72,  5.0)),
    ("VBR", (0.15, 0.20,  8.0), (0.22, 0.17,  8.0)),
    ("VAB", (0.25, 0.10, 32.0), (0.27, 0.10, 29.0)),
    ("RFB", (0.34, 0.77,  7.0), (0.38, 0.71, 14.0)),
    ("VBZ", (0.45, 0.55,  6.0), (0.48, 0.53, 12.0)),
    ("RS6", (0.72, 0.43,  3.0), (0.75, 0.41,  6.0)),
    ("V11", (0.77, 0.75,  4.0), (0.71, 0.73,  4.0)),
    ("V12", (0.78, 1.10,  1.1), (0.74, 1.15,  1.1)),
)

# Bought in December 2025, so they have a position and a size but no PY twin.
# That absence is the point of the message and the reason they are drawn in the
# highlight blue rather than the actual grey.
C10D_ACQUIRED = (
    ("RFA", (0.06, 0.30, 3.0)),
    ("VAA", (0.10, 0.62, 1.0)),
)

# The scenario the message is about. Named once rather than spelled "ACQ" at a
# dozen call sites, because both renderers have to agree what it is.
C10D_ACQUISITION = "ACQ"

# The three names the reference does not print. Where a pair of bubbles sits
# close enough that the two names would run into each other, it prints one and
# keeps the prior year's - which in all three cases is the left-hand bubble, so
# the pair still reads left to right.
#
# Transcribed rather than computed, like C02A's suppressed figures. Where the
# reference draws the line is far too fine to make a rule of: RS6's two names
# would sit 32px apart and one of them goes, VBZ's would sit 33px apart and both
# stay. Every value label is printed - it is only the names that give way.
C10D_UNLABELLED = (("RS6", "AC"), ("VBR", "AC"), ("VAB", "AC"))

# Each bubble's radius in the 1280x720 reference render, in pixels, from the
# free circle fit. Measurements, not settings: the renderers compute their own
# radii from the area law, and these exist only to be checked against.
C10D_MEASURED_RADII = (
    ("RS1", "PY", 28.50), ("RS1", "AC", 43.30),
    ("VBN", "PY", 19.52), ("VBN", "AC", 31.09),
    ("VBR", "PY", 40.30), ("VBR", "AC", 42.72),
    ("VAB", "PY", 83.28), ("VAB", "AC", 78.72),
    ("RFB", "PY", 37.93), ("RFB", "AC", 54.31),
    ("VBZ", "PY", 32.80), ("VBZ", "AC", 49.92),
    ("RS6", "PY", 24.40), ("RS6", "AC", 35.01),
    ("V11", "PY", 28.50), ("V11", "AC", 28.35),
    ("V12", "PY", 14.17), ("V12", "AC", 13.85),
    ("RFA", "ACQ", 24.15), ("VAA", "ACQ", 13.15),
)


def _c10d_points() -> tuple["Point", ...]:
    """Every bubble, in the order they are painted.

    Order matters here in a way it does not elsewhere: the bubbles overlap and
    are opaque, so whatever is drawn last wins the middle of the pile. Prior
    year first, then actual, then the acquisitions - so the current position of
    a unit is never hidden behind where it used to be. That is a notation
    decision rather than a drawing convenience: VAB's actual (29) is *smaller*
    than its prior year (32), so painting biggest-first - the usual bubble-chart
    habit - would bury this year's position under last year's.
    """
    out: list[Point] = []
    for name, py, _ac in C10D_SBUS:
        out.append(Point(name, "PY", *py))
    for name, _py, ac in C10D_SBUS:
        out.append(Point(name, "AC", *ac))
    for name, ac in C10D_ACQUIRED:
        out.append(Point(name, C10D_ACQUISITION, *ac))
    return tuple(out)


C10D_POINTS = _c10d_points()

C10D = Template(
    id="C10",
    variant="D",
    kind="bubble chart",
    title=TitleBlock(
        entity="Alpha Corp., Strategic business units",
        measure="Product market portfolio",
        # No unit on the subject line. Three measures are plotted and they are
        # in three different units, so no single one can head the chart; each
        # is named where it is used - two axis titles and the size legend.
        unit="",
        period="2025",
        message=("The two SBUs acquired in December 2025 "
                 "are positioned in little attractive markets"),
    ),
    # An XY chart has no category axis. The entity names are still categories in
    # the ordinary sense - they are what a row of the source table is - and each
    # of them carries an actual reading.
    categories=(tuple(n for n, _p, _a in C10D_SBUS)
                + tuple(n for n, _a in C10D_ACQUIRED)),
    category_scenarios=("AC",) * (len(C10D_SBUS) + len(C10D_ACQUIRED)),
    tiers=(),
    points=C10D_POINTS,
    axes=C10D_AXES,
    # What a bubble's area counts. The legend prints it as its heading, which is
    # where the chart's third unit is stated.
    size_legend=("Net sales", "mEUR"),
    source_ref="template-refs/C10_10D.png",
    notes=(
        "Nine business units drawn twice and two December acquisitions drawn "
        "once; the missing prior year is what the blue says.",
        "Bubble area is proportional to net sales, never radius - the eye "
        "reads area, and a radius-proportional mark overstates by the square.",
        "Prior year is painted first so an actual is never hidden behind where "
        "the unit used to be, even where the actual is the smaller bubble.",
    ),
)


def _fit_line(xs: Sequence[float], ys: Sequence[float]) -> tuple[float, float]:
    """Ordinary least squares slope and intercept. Plain arithmetic, no numpy."""
    n = len(xs)
    mx, my = sum(xs) / n, sum(ys) / n
    sxy = sum((x - mx) * (y - my) for x, y in zip(xs, ys))
    sxx = sum((x - mx) ** 2 for x in xs)
    slope = sxy / sxx
    return slope, my - slope * mx


def _check_c10d(check) -> None:
    """Tie-outs for C10D.

    Nothing in a portfolio chart is derived, so most of these check the *claims*
    rather than the arithmetic: that the sentence over the chart is true of the
    points under it, that every point is on the page, and that the paired units
    really are paired. A portfolio chart whose message its own points do not
    support is the characteristic failure of this template, and it is invisible
    to any check that only adds numbers up.

    The last one is the transcription tie-out, and it is the strict one.
    """
    axis_x, axis_y = C10D_AXES

    off = [f"{p.entity} {p.scenario} at ({p.x:.2f}, {p.y:.2f})"
           for p in C10D_POINTS
           if not (axis_x.minimum <= p.x <= axis_x.maximum
                   and axis_y.minimum <= p.y <= axis_y.maximum)]
    check("every bubble centre falls inside the axes as drawn",
          not off,
          f"all {len(C10D_POINTS)} within attractiveness "
          f"{axis_x.minimum:.2f}..{axis_x.maximum:.2f} and share "
          f"{axis_y.minimum:.2f}..{axis_y.maximum:.2f}"
          if not off else "; ".join(off))

    # The message. Both acquisitions have to sit at the unattractive end, and
    # "little attractive" has to mean something a reader can check - so the test
    # is against the units already in the portfolio, not against a threshold
    # picked to make the sentence come out true.
    acquired = {p.entity: p for p in C10D_POINTS
                if p.scenario == C10D_ACQUISITION}
    incumbent = min(p.x for p in C10D_POINTS if p.scenario == "AC")
    worst = max(p.x for p in acquired.values())
    check(
        "the two acquisitions really are in the least attractive markets",
        worst <= incumbent,
        f"both bought units sit at attractiveness {worst:.2f} or below, "
        f"against {incumbent:.2f} for the least attractive unit the group "
        f"already had - so the sentence is true of the picture",
    )

    check(
        "an acquisition has an actual and no prior year",
        len(acquired) == 2
        and not any(p.entity in acquired and p.scenario == "PY"
                    for p in C10D_POINTS),
        f"{', '.join(sorted(acquired))} are drawn once each; a unit bought in "
        f"December has no prior year inside this group to be compared with",
    )

    paired = [n for n, _p, _a in C10D_SBUS]
    counts = {n: sum(1 for p in C10D_POINTS if p.entity == n) for n in paired}
    check("every unit the group already owned is drawn twice",
          all(c == 2 for c in counts.values()),
          f"all {len(paired)} carry a prior year and an actual"
          if all(c == 2 for c in counts.values())
          else "; ".join(f"{n} appears {c} time(s)"
                         for n, c in counts.items() if c != 2))

    # Relative market share is sales over the largest competitor's, so 1.00 is
    # the line between leading a market and following in it. Which side of it a
    # unit sits on is the reason the axis runs to 1.25 rather than stopping at
    # 1.00: an axis that stopped there would clip the leaders off the page.
    leaders_py = sorted(n for n, py, _ac in C10D_SBUS if py[1] > 1.0)
    leaders_ac = sorted(n for n, _py, ac in C10D_SBUS if ac[1] > 1.0)
    check(
        "the share axis runs past 1.00 because units sit above it",
        max(p.y for p in C10D_POINTS) > 1.0 and axis_y.maximum > 1.0,
        f"market leaders last year: {', '.join(leaders_py) or 'none'}; this "
        f"year: {', '.join(leaders_ac) or 'none'} - RS1 fell to exactly 1.00, "
        f"which is the line itself",
    )

    grew = [n for n, py, ac in C10D_SBUS if ac[2] > py[2]]
    flat = [n for n, py, ac in C10D_SBUS if ac[2] == py[2]]
    fell = [n for n, py, ac in C10D_SBUS if ac[2] < py[2]]
    total_py = sum(py[2] for _n, py, _ac in C10D_SBUS)
    total_ac = sum(ac[2] for _n, _py, ac in C10D_SBUS)
    bought = sum(ac[2] for _n, ac in C10D_ACQUIRED)
    check(
        "net sales move the way the pairs of bubbles say they do",
        len(grew) + len(flat) + len(fell) == len(C10D_SBUS) and len(fell) == 1,
        f"{len(grew)} units grew ({', '.join(grew)}), {len(flat)} held "
        f"({', '.join(flat)}) and {', '.join(fell)} alone fell - "
        f"{total_py:.1f} mEUR last year against {total_ac:.1f} now, plus "
        f"{bought:.1f} bought in December for {total_ac + bought:.1f} in all",
    )

    # The transcription tie-out. Area is proportional to value, so a drawn
    # radius is k*sqrt(value) less a constant inset - the white outline each
    # bubble carries, which a measurement of the visible disc loses at both
    # edges. Twenty measurements against nine distinct values, fitted with two
    # free parameters, leaves eighteen degrees of freedom: if a value has been
    # read onto the wrong bubble the line stops fitting. Misreading 14.0 as 4.0
    # takes the error to 5.7px and 32.0 as 3.0 takes it to 12.8px, against 0.7
    # as transcribed. What it cannot see is 9 against 8 - two values one step
    # apart draw radii 2px apart, and that is inside the measurement noise.
    sizes = {(p.entity, p.scenario): p.size for p in C10D_POINTS}
    missing = [f"{n} {s}" for n, s, _r in C10D_MEASURED_RADII
               if (n, s) not in sizes]
    if missing:
        # A bubble was measured off the reference and is not in the data. Say
        # so; do not fall over on a KeyError three frames down.
        check("one area law explains all twenty measured bubbles", False,
              f"measured but not drawn: {', '.join(missing)}")
        return
    roots = [sizes[(n, s)] ** 0.5 for n, s, _r in C10D_MEASURED_RADII]
    radii = [r for _n, _s, r in C10D_MEASURED_RADII]
    k, inset = _fit_line(roots, radii)
    errors = [abs(r - (k * s + inset)) for s, r in zip(roots, radii)]
    rms = (sum(e * e for e in errors) / len(errors)) ** 0.5
    check(
        "one area law explains all twenty measured bubbles",
        rms < 1.2 and abs(k - 15.0) < 0.5,
        f"radius = {k:.2f}*sqrt(net sales) - {abs(inset):.2f}px, fitting all "
        f"{len(radii)} to {rms:.2f}px rms and {max(errors):.2f}px at worst "
        f"(VBZ's prior year and VBR's actual, the two most heavily overlapped, "
        f"where the visible arc is shortest)",
    )


# --------------------------------------------------------------------------- #
# C09C - scattergram, Alpha Corp. paper division 2025
# --------------------------------------------------------------------------- #
#
# A hundred and forty-nine products, three product lines, and two measures that
# multiply into a third. Margin runs across, net sales runs up, and gross profit
# is their product - which is why the reference draws curves rather than a line:
# every point on sales x margin = 3 mUSD earns the same gross profit, and the
# curve through them is what turns a cloud into a statement.
#
# All 149 positions were measured off the reference render rather than invented.
# An earlier note in decisions.md said this variant's points would have to be
# synthesized because none of C09's variants labels them; that was wrong, and
# usefully so. The values are not printed, but they are *drawn*, and a drawn
# position is a measurement: each marker is an 18px disc, so classifying every
# pixel by hue and then fitting disc centres recovers all three coordinates to
# about a twentieth of a unit. Nothing here is made up.
#
# Two things that measurement had to get past, both worth knowing:
#
#   The markers are translucent. A mask keyed on the exact fill colour finds
#   them over paper and loses them over each other, over the shaded segment and
#   over the curves - which is most of the interesting part of the chart. Hue
#   survives all three, because drawing orange over grey changes how light it is
#   and not what colour it is.
#
#   Where markers overlap, only the topmost is a whole disc. The rest were
#   recovered the way C10's occluded bubbles were: a hidden marker is still a
#   full disc inside (its own colour) union (the markers already found).
#
# The count is one short of the message, and that is the finding rather than a
# failure - see _check_c09c.

C09C_LINES = ("RX", "X2", "VA")

# Which accent each product line takes, as an index into the palette's accent
# ramp. Held here rather than in either renderer because both draw it, and
# because it is not the legend's order: the ramp was sampled from this very
# chart, so this is the order the sampler walked in.
C09C_ACCENT = {"RX": 0, "VA": 1, "X2": 2}

C09C_AXES = (
    Axis("Margin", 0.0, 35.0, 5.0, "{:.0f}", unit="% of net sales"),
    Axis("Net sales", 0.0, 35.0, 5.0, "{:.0f}", unit="mUSD"),
)

# Gross profit is net sales times margin, so a constant gross profit is a
# hyperbola. Three of them are drawn, and the top one is also the boundary of
# the shaded segment the message is about.
C09C_ISO_PROFIT = (1.0, 2.0, 3.0)
C09C_SEGMENT = 3.0

# What the message claims, held as a number rather than left inside the
# sentence, because a tie-out has to be able to read it.
C09C_MESSAGE_LINE = "VA"
C09C_MESSAGE_COUNT = 45

# Product line, margin in % of net sales, net sales in mUSD, and - for the ten
# the reference names - the product. Measured; see the note above.
#
# One product sits centred on the top rule and measures 35.17, which is
# 0.17 above an axis that stops at 35. The rule is two pixels thick and drawn
# through the marker, which is enough to pull a disc fit that far; it is snapped
# back to 35.00 rather than left to fail the range check on a rounding of the
# reference's own gridline.
C09C_PRODUCTS = (
    ("RX",  2.14,  3.53, "RX-180"),
    ("RX",  7.49, 22.72),
    ("RX",  7.93, 16.05),
    ("RX",  8.53, 27.90),
    ("RX",  8.97, 25.27),
    ("RX",  9.00, 20.02),
    ("RX",  9.04, 24.67),
    ("RX",  9.20, 24.06),
    ("RX",  9.51, 23.86),
    ("RX", 10.42, 15.99),
    ("RX", 10.75, 19.15),
    ("RX", 11.09, 28.71),
    ("RX", 11.32,  9.86),
    ("RX", 11.83, 16.73),
    ("RX", 12.33, 28.17),
    ("RX", 12.40, 26.49),
    ("RX", 12.44, 14.50),
    ("RX", 12.47, 23.66),
    ("RX", 12.91, 28.30),
    ("RX", 13.81, 21.44),
    ("RX", 14.25, 19.62),
    ("RX", 14.62, 24.87),
    ("RX", 14.72, 25.48),
    ("RX", 14.82, 21.30),
    ("RX", 15.06, 10.26),
    ("RX", 15.30, 18.75),
    ("RX", 15.40, 21.98),
    ("RX", 15.40, 13.09),
    ("RX", 15.53, 19.35),
    ("RX", 15.67, 19.96),
    ("RX", 15.70, 21.71),
    ("RX", 15.94, 23.73),
    ("RX", 16.04, 14.64),
    ("RX", 16.27, 25.61),
    ("RX", 16.27, 22.04),
    ("RX", 16.34, 19.89),
    ("RX", 16.64, 19.55),
    ("RX", 16.74, 18.95),
    ("RX", 19.27, 12.42),
    ("RX", 19.27, 10.80),
    ("RX", 20.07,  9.86),
    ("RX", 22.06,  8.72),
    ("RX", 22.36, 20.63),
    ("RX", 22.77,  8.99),
    ("RX", 25.06, 21.17, "RX-6"),
    ("RX", 25.86, 24.60, "RX-4"),
    ("RX", 31.25, 34.90, "RX-2000"),
    ("RX", 33.06,  7.77),
    ("X2",  2.17, 16.12),
    ("X2",  2.71, 17.13),
    ("X2",  3.79, 15.31),
    ("X2",  3.85, 23.32),
    ("X2",  3.89, 10.94),
    ("X2",  4.29, 21.71),
    ("X2",  4.32, 20.09),
    ("X2",  4.32, 18.75),
    ("X2",  4.53,  9.19),
    ("X2",  4.63, 17.67),
    ("X2",  4.93, 15.99),
    ("X2",  5.10,  7.44),
    ("X2",  5.33, 21.77),
    ("X2",  5.64, 21.17),
    ("X2",  5.84, 11.81),
    ("X2",  6.18, 17.06),
    ("X2",  6.28, 19.62),
    ("X2",  6.34,  8.25),
    ("X2",  6.48, 22.38),
    ("X2",  6.48, 16.32),
    ("X2",  6.55, 15.11),
    ("X2",  6.71, 10.53),
    ("X2",  7.82, 21.71),
    ("X2",  7.99,  9.05),
    ("X2",  8.03,  5.15),
    ("X2",  8.50,  5.22),
    ("X2",  8.53, 13.09),
    ("X2",  8.70, 10.13),
    ("X2",  8.77,  8.78),
    ("X2",  8.97, 12.42),
    ("X2", 10.31, 12.22),
    ("X2", 11.22,  5.28),
    ("X2", 11.26, 19.55),
    ("X2", 14.69,  3.40),
    ("X2", 15.50,  2.86),
    ("X2", 15.97,  2.73),
    ("X2", 21.02,  2.66),
    ("X2", 21.79,  2.66),
    ("X2", 22.36,  3.06),
    ("X2", 23.44,  2.32),
    ("X2", 24.18,  2.59),
    ("X2", 24.72,  2.32),
    ("X2", 27.04, 26.82, "X2-122"),
    ("X2", 31.28, 30.12, "X2-201"),
    ("X2", 31.38,  1.38),
    ("X2", 33.50, 33.01, "X2-200"),
    ("X2", 34.85,  1.51),
    ("VA",  2.17,  8.04, "VA-D2"),
    ("VA",  4.02, 32.27),
    ("VA",  5.33, 16.93),
    ("VA",  5.40,  3.13, "VA-3PO"),
    ("VA",  7.45, 35.00),
    ("VA", 11.39, 21.77),
    ("VA", 13.81, 24.20),
    ("VA", 14.19, 27.90),
    ("VA", 14.56, 26.96),
    ("VA", 15.09, 26.35),
    ("VA", 15.50, 28.57),
    ("VA", 15.50, 25.14),
    ("VA", 15.53, 31.00),
    ("VA", 15.97,  7.98),
    ("VA", 16.00, 32.00),
    ("VA", 16.47, 34.02),
    ("VA", 16.54, 27.36),
    ("VA", 16.64, 25.00),
    ("VA", 17.28, 28.37),
    ("VA", 17.28, 27.76),
    ("VA", 17.69, 10.94),
    ("VA", 17.99, 15.25),
    ("VA", 18.32, 33.49),
    ("VA", 18.46, 20.43),
    ("VA", 18.53, 31.26),
    ("VA", 18.76, 24.33),
    ("VA", 19.54, 34.09),
    ("VA", 19.54, 24.67),
    ("VA", 19.60, 34.70),
    ("VA", 19.74, 19.49),
    ("VA", 19.84, 24.47),
    ("VA", 20.14, 18.88),
    ("VA", 20.38, 20.76),
    ("VA", 20.81, 25.68),
    ("VA", 20.92, 22.78),
    ("VA", 20.95, 32.48),
    ("VA", 22.63, 33.01),
    ("VA", 22.73, 15.78),
    ("VA", 22.87, 20.36),
    ("VA", 23.14, 23.93),
    ("VA", 24.95, 18.21),
    ("VA", 24.99, 20.43),
    ("VA", 25.22, 16.05),
    ("VA", 25.69,  8.99),
    ("VA", 26.23, 17.94),
    ("VA", 26.23, 17.33),
    ("VA", 26.54, 17.33),
    ("VA", 26.57, 17.94),
    ("VA", 28.05, 16.59),
    ("VA", 29.16, 32.48, "VA-19"),
    ("VA", 29.26, 14.10),
    ("VA", 29.36, 15.45),
    ("VA", 29.36, 14.84),
    ("VA", 34.85, 15.04),
)


def _c09c_points() -> tuple["Point", ...]:
    """Every product as a point, in the order they are painted.

    Product line by product line rather than interleaved, because a line is a
    *series* on both engines - one fill, one legend entry - and because what is
    drawn last sits on top. RX, X2, VA is the reference's own legend order and
    therefore its paint order too.

    ``scenario`` is AC for all of them and ``group`` carries the product line.
    That split is the whole notational point of this template: a product line is
    not a scenario. Every one of these 149 products is an actual, measured the
    same way in the same year, and colouring them from the accent ramp rather
    than the scenario greys is what stops the reader taking orange to mean
    something about how the number came to exist.
    """
    order = {name: i for i, name in enumerate(C09C_LINES)}
    rows = sorted(C09C_PRODUCTS, key=lambda r: order[r[0]])
    return tuple(
        Point(row[3] if len(row) > 3 else "", "AC", row[1], row[2], group=row[0])
        for row in rows
    )


C09C_POINTS = _c09c_points()


def c09c_gross_profit(point: "Point") -> float:
    """Gross profit in mUSD: net sales times margin.

    One function, read by the tie-outs and by both renderers, so the curves
    drawn on the page and the counts asserted about them cannot come from two
    different definitions of the same quantity.
    """
    return point.y * point.x / 100.0


C09C = Template(
    id="C09",
    variant="C",
    kind="scattergram",
    title=TitleBlock(
        entity="Alpha Corp., Paper division",
        measure="Net sales",
        unit="mUSD, margin in %",
        period="2025",
        message=("In 2025 we had 45 products of the product line VA "
                 "in the gross profit segment of 3 mUSD and above"),
        subject=(("Net sales", "bold"), (" in mUSD, ", "normal"),
                 ("margin", "bold"), (" in %", "normal")),
    ),
    # An XY chart has no category axis. Each product is a category in the
    # ordinary sense; most of them have no name because the reference prints
    # none, and inventing 139 product codes would be inventing data.
    categories=tuple(p.entity for p in C09C_POINTS),
    category_scenarios=("AC",) * len(C09C_POINTS),
    tiers=(),
    points=C09C_POINTS,
    axes=C09C_AXES,
    source_ref="template-refs/C09_09C.png",
    notes=(
        "149 products measured off the reference render; none of C09's "
        "variants prints its values, but a drawn position is a measurement.",
        "Colour carries the product line, which is a category and not a "
        "scenario - the accent ramp, never the scenario greys.",
        "Gross profit is the product of the two axes, so a constant gross "
        "profit is a curve; the top one bounds the segment the message names.",
        "The chart draws 44 VA products in that segment where the message "
        "says 45. One of them is under another and cannot be seen.",
    ),
)


def _check_c09c(check) -> None:
    """Tie-outs for C09C.

    A scattergram states no arithmetic between its two axes, but it does state
    one *about* them - gross profit is their product - and the whole chart is
    built on that: the curves, the shaded segment and the sentence over the top
    all depend on it. So these checks are mostly about the third measure nobody
    plots directly, and about whether the message is true of the picture.
    """
    axis_x, axis_y = C09C_AXES

    off = [f"{p.group} at ({p.x:.2f}, {p.y:.2f})" for p in C09C_POINTS
           if not (axis_x.minimum <= p.x <= axis_x.maximum
                   and axis_y.minimum <= p.y <= axis_y.maximum)]
    check("every product falls inside the axes as drawn", not off,
          f"all {len(C09C_POINTS)} within margin "
          f"{axis_x.minimum:.0f}..{axis_x.maximum:.0f}% and net sales "
          f"{axis_y.minimum:.0f}..{axis_y.maximum:.0f} mUSD"
          if not off else "; ".join(off))

    counts = {line: sum(1 for p in C09C_POINTS if p.group == line)
              for line in C09C_LINES}
    check(
        "every product belongs to one of the three lines the legend names",
        sum(counts.values()) == len(C09C_POINTS)
        and not any(p.group not in C09C_LINES for p in C09C_POINTS),
        ", ".join(f"{line} {n}" for line, n in counts.items())
        + f" - {sum(counts.values())} products in all",
    )

    # The message, and the one place the chart and its own sentence disagree.
    #
    # The count is 44, not the 45 the reference prints. That is not a
    # measurement failure: the detection is stable at 44 across every threshold,
    # every marker radius and every occlusion setting tried, and a search for
    # unexplained fill turned up nothing marker-sized. The explanation is the
    # honest one, and it is the lesson of the template: two products of the same
    # line at nearly the same margin and the same net sales draw one on top of
    # the other, and where the colours match there is no edge left to see. A
    # scattergram cannot be counted by looking at it.
    #
    # So the recreation prints the reference's sentence unaltered - the company
    # had 45 - and this check asserts the size of the gap rather than papering
    # over it. Inventing a 150th product to sit under an existing one would make
    # the arithmetic agree and the transcription false.
    segment = [p for p in C09C_POINTS
               if p.group == C09C_MESSAGE_LINE
               and c09c_gross_profit(p) >= C09C_SEGMENT]
    gap = C09C_MESSAGE_COUNT - len(segment)
    check(
        "the segment the message counts, against the count it states",
        gap == 1,
        f"{len(segment)} VA products are drawn at {C09C_SEGMENT:.0f} mUSD gross "
        f"profit or more, against the {C09C_MESSAGE_COUNT} the message states - "
        + ("one is behind another of its own colour and leaves no edge to find"
           if gap == 1 else
           f"a gap of {gap}, where the reference's own picture is one short of "
           f"its own sentence and no more"),
    )

    # The three named at the low end are exactly the three lowest gross profits
    # on the page, which is what makes naming them mean something.
    named = [p for p in C09C_POINTS if p.entity]
    lowest = sorted(C09C_POINTS, key=c09c_gross_profit)[:3]
    check(
        "the three products named at the bottom are the three least profitable",
        {p.entity for p in lowest} == {"RX-180", "VA-3PO", "VA-D2"},
        ", ".join(f"{p.entity} at {c09c_gross_profit(p):.2f} mUSD"
                  for p in lowest)
        + " - the three lowest gross profits of all 149",
    )

    # And the names have to agree with the colours, which is a transcription
    # check: a label matched to the wrong marker shows up here at once.
    wrong = [f"{p.entity} is drawn as {p.group}" for p in named
             if not p.entity.startswith(p.group + "-")]
    check(
        "every named product is drawn in its own product line's colour",
        not wrong and len(named) == 10,
        f"all {len(named)} names agree with the marker they sit over"
        if not wrong else "; ".join(wrong),
    )

    # The curves. Gross profit is the product of the axes, so the curve for a
    # given profit has to pass through the right height at the right margin -
    # and the check is worth having because the obvious mistake, drawing the
    # curve as a straight line between its endpoints, looks plausible.
    worst = 0.0
    for level in C09C_ISO_PROFIT:
        for margin in (5.0, 15.0, 25.0, 35.0):
            sales = level * 100.0 / margin
            worst = max(worst, abs(sales * margin / 100.0 - level))
    check(
        "each iso-profit curve is the locus of one gross profit",
        worst < 1e-9,
        f"sales = 100 x profit / margin reproduces {', '.join(f'{c:.0f}' for c in C09C_ISO_PROFIT)}"
        f" mUSD exactly at every margin tested; at 35% the three curves sit at "
        + ", ".join(f"{c * 100 / 35:.2f}" for c in C09C_ISO_PROFIT) + " mUSD",
    )

    # The segment boundary is one of the drawn curves, not a fourth line. If it
    # ever stops being, the shading and the sentence part company.
    check(
        "the shaded segment is bounded by a curve the chart already draws",
        C09C_SEGMENT in C09C_ISO_PROFIT,
        f"the {C09C_SEGMENT:.0f} mUSD curve is both the top iso-profit line and "
        f"the edge of the segment, so the reader is not asked to trust an "
        f"unlabelled boundary",
    )


# --------------------------------------------------------------------------- #
# C11A - ROI tree, Alpha Software Corporation Germany 2021..2027
# --------------------------------------------------------------------------- #

C11A_YEARS = ("2021", "2022", "2023", "2024", "2025", "2026", "2027")
# Five actual years and two planned. The switch is drawn as a vertical rule
# between 2025 and 2026 in every one of the six boxes, which is what lets a
# reader carry the AC/PL boundary across the tree without re-reading six axes.
C11A_SCENARIOS = ("AC", "AC", "AC", "AC", "AC", "PL", "PL")

# The three base measures - the right-hand column of the tree. These are the
# typed inputs; everything else on the page is arithmetic between them.
C11A_RETURN = (5.0, -1.8, -3.4558, 3.1, 4.1, 4.7, 5.5)
C11A_NET_SALES = (26.1, 25.7, 26.4, 22.1, 23.8, 23.7, 27.7)
C11A_CAPITAL = (20.1, 18.5, 19.4, 21.1, 22.9, 24.5, 24.5)

# 2023's return is carried to four decimals, and the reason is worth stating.
# Every figure IBCS prints here is rounded to a tenth, and taking a quotient of
# two rounded figures compounds the two roundings - the same trap C03A's PY
# derivation documents. Typing -3.5 makes return on sales print -13.3 where the
# original prints -13.1, and ROI print -18.0 where the original prints -17.8.
#
# The precise value is not invented, it is *recovered*, and by two independent
# routes that agree: the printed -13.1% over net sales of 26.4 gives -3.4584,
# and the printed -17.8% over capital of 19.4 gives -3.4532. Two separate
# printed ratios, agreeing to 0.005 kEUR, whose mean rounds to -3.5 - which is
# what the return box itself prints. With it, all twenty-one derived labels
# reproduce IBCS's own print exactly; without it, two do not.
C11A_RETURN_RECOVERY = (-13.1 / 100 * 26.4, -17.8 / 100 * 19.4)

# What the six boxes print, transcribed from the reference. Geometry is drawn
# from the derived values; these are only what the labels say.
C11A_PRINTED_RETURN = (5.0, -1.8, -3.5, 3.1, 4.1, 4.7, 5.5)
C11A_PRINTED_NET_SALES = (26.1, 25.7, 26.4, 22.1, 23.8, 23.7, 27.7)
C11A_PRINTED_CAPITAL = (20.1, 18.5, 19.4, 21.1, 22.9, 24.5, 24.5)
C11A_PRINTED_ROS = (19.2, -7.0, -13.1, 14.0, 17.2, 19.8, 19.9)
C11A_PRINTED_TURNOVER = (1.3, 1.4, 1.4, 1.0, 1.0, 1.0, 1.1)
C11A_PRINTED_ROI = (24.9, -9.7, -17.8, 14.7, 17.9, 19.2, 22.4)


# Every bar's drawn extent, measured off the reference render in pixels from
# each box's own zero line. Positive is up. These are not an input to the
# drawing - they are how the shared-scale claim is *checked*: if each unit
# group really is one ruler, then one number per group reproduces all 35 of
# them, and the residual says how well.
C11A_MEASURED_EXTENTS = {
    "roi": (132, -52, -95, 78, 95, 102, 119),
    "ros": (102, -37, -70, 75, 92, 106, 106),
    "turnover": (173, 185, 182, 140, 139, 129, 151),
    "return": (21, -8, -15, 13, 18, 20, 24),
    "net_sales": (112, 110, 113, 95, 102, 102, 119),
    "capital": (86, 79, 83, 90, 98, 104, 104),
}

# Fitted from those measurements, one per unit group, and used by both
# renderers as the scale. A group's ruler is a single number: px per percent,
# px per kEUR, px per turn.
C11A_SCALE_PX = {"percent": 5.3231, "kEUR": 4.2793, "turnover": 133.4735}


def _c11a_quotient(num: Sequence[float], den: Sequence[float],
                   scale: float = 1.0) -> tuple[float, ...]:
    return tuple(n / d * scale for n, d in zip(num, den))


# The three derived measures. Module constants rather than something computed
# inside a renderer, so that both engines, the tie-outs and the worksheet
# formulas are all talking about the same numbers.
C11A_ROS = _c11a_quotient(C11A_RETURN, C11A_NET_SALES, 100.0)
C11A_TURNOVER = _c11a_quotient(C11A_NET_SALES, C11A_CAPITAL)
C11A_ROI = _c11a_quotient(C11A_RETURN, C11A_CAPITAL, 100.0)


def _c11a_tier(key: str, label: str, unit: str, values: Sequence[float],
               printed: Sequence[float], fmt: str) -> Tier:
    """One box of the tree, as a measure tier with an AC and a PL series.

    The scenario lives in the series, not in the box, because a single box runs
    actual then plan - which is exactly the case ``_split`` exists for. Passing
    ``C11A_SCENARIOS`` explicitly matters: the default is C03A's October
    switch, and C11 turns to plan after 2025.
    """
    return Tier(
        key=key,
        label=label,
        kind="measure",
        number_format=fmt,
        printed=printed,
        block=unit,
        series=(
            Series("AC", _split(values, "AC", C11A_SCENARIOS)),
            Series("PL", _split(values, "PL", C11A_SCENARIOS)),
        ),
    )


C11A_TREE = TreeSpec(
    nodes=(
        # Column 0 is the result the page is about, and the tree reads right to
        # left: the base measures on the right combine into the two ratios, and
        # those two combine into return on investment.
        TreeNode(key="roi", column=0, scale_group="percent", unit="%"),
        TreeNode(key="ros", column=1, scale_group="percent", unit="%"),
        TreeNode(key="turnover", column=1, scale_group="turnover"),
        TreeNode(key="return", column=2, scale_group="kEUR", unit="kEUR"),
        TreeNode(key="net_sales", column=2, scale_group="kEUR", unit="kEUR"),
        TreeNode(key="capital", column=2, scale_group="kEUR", unit="kEUR"),
    ),
    links=(
        TreeLink(operator=":", left="return", right="net_sales", result="ros"),
        TreeLink(operator=":", left="net_sales", right="capital",
                 result="turnover"),
        TreeLink(operator="x", left="ros", right="turnover", result="roi"),
    ),
    printed_categories=(0, 3, 4, 5, 6),
)


C11A = Template(
    id="C11",
    variant="A",
    kind="tree chart",
    title=TitleBlock(
        entity="Alpha Software Corporation, Germany",
        measure="ROI tree",
        # No unit on the subject line: the tree carries three of them - a
        # percentage, a currency and a dimensionless ratio - so no single one
        # can head the page. Each box names its own, which is also what makes
        # the scale groups legible.
        unit="",
        period="2021..2027",
        # Printed exactly as the reference does, including "despite of" and
        # including the year. IBCS's own message names 2024, but 19.2% is 2026:
        # 2024's ROI is 14.7, and the highlight oval on the page sits over
        # 2026. Recorded as a tie-out rather than silently corrected - the same
        # treatment as T03A's variance and C09C's product count.
        message=("We plan to achieve a ROI of 19,2% in 2024 "
                 "despite of increasing invested capital"),
    ),
    categories=C11A_YEARS,
    category_scenarios=C11A_SCENARIOS,
    tiers=(
        _c11a_tier("roi", "Return on investment", "percent",
                   C11A_ROI, C11A_PRINTED_ROI, "{:,.1f}"),
        _c11a_tier("ros", "Return on sales", "percent",
                   C11A_ROS, C11A_PRINTED_ROS, "{:,.1f}"),
        _c11a_tier("turnover", "Capital turnover", "turnover",
                   C11A_TURNOVER, C11A_PRINTED_TURNOVER, "{:,.1f}"),
        _c11a_tier("return", "Return", "kEUR",
                   C11A_RETURN, C11A_PRINTED_RETURN, "{:,.1f}"),
        _c11a_tier("net_sales", "Net sales", "kEUR",
                   C11A_NET_SALES, C11A_PRINTED_NET_SALES, "{:,.1f}"),
        _c11a_tier("capital", "Invested capital", "kEUR",
                   C11A_CAPITAL, C11A_PRINTED_CAPITAL, "{:,.1f}"),
    ),
    tree=C11A_TREE,
    # The oval marks the planned ROI the message is about - 2026, the first
    # planned year, not the 2024 the sentence names.
    annotations=(Annotation(kind="oval", target=("roi", 5)),),
    source_ref="template-refs/C11_11A.png",
    notes=(
        "Six small charts and the arithmetic between them: return over net "
        "sales is return on sales, net sales over invested capital is capital "
        "turnover, and their product is return on investment.",
        "Boxes measuring the same unit share a scale. Return peaks at 5.5 "
        "against net sales' 27.7 and both are drawn at the same px-per-kEUR - "
        "the disparity is the point, and rescaling the return box to fill its "
        "own height would hide it.",
        "The box heights differ because each is sized to its own range; the "
        "scale does not, which is the rule that makes the three currency "
        "boxes comparable.",
    ),
)


def _check_c11a(check) -> None:
    """Tie-outs for C11A.

    A tree is a set of claims about arithmetic, so most of these evaluate the
    links rather than compare transcriptions. The one that matters most is that
    the *derived* ratios reproduce the labels IBCS printed: that is what says
    the three typed series are the right three numbers, and it is the check
    that forced 2023's return to be recovered rather than read.
    """
    n = len(C11A_YEARS)

    check(
        "every series runs the full seven years",
        all(len(s) == n for s in (C11A_RETURN, C11A_NET_SALES, C11A_CAPITAL,
                                  C11A_ROS, C11A_TURNOVER, C11A_ROI)),
        "seven periods, 2021..2027, in all six boxes",
    )

    # The two independent recoveries of 2023's return.
    a, b = C11A_RETURN_RECOVERY
    check(
        "2023's return is recovered from two printed ratios that agree",
        abs(a - b) < 0.01 and round((a + b) / 2, 1) == C11A_PRINTED_RETURN[2],
        f"net sales route {a:.4f}, capital route {b:.4f}, differ by "
        f"{abs(a - b):.4f}; their mean {(a + b) / 2:.4f} prints "
        f"{(a + b) / 2:.1f} and the return box prints "
        f"{C11A_PRINTED_RETURN[2]:.1f}",
    )

    # Every derived label reproduces IBCS's print. Twenty-one of them.
    mismatches = []
    for name, derived, printed in (
        ("return on sales", C11A_ROS, C11A_PRINTED_ROS),
        ("capital turnover", C11A_TURNOVER, C11A_PRINTED_TURNOVER),
        ("return on investment", C11A_ROI, C11A_PRINTED_ROI),
    ):
        for i, year in enumerate(C11A_YEARS):
            if round(derived[i], 1) != printed[i]:
                mismatches.append(
                    f"{name} {year}: derived {derived[i]:.3f} prints "
                    f"{round(derived[i], 1)}, reference prints {printed[i]}"
                )
    check(
        "all 21 derived ratio labels reproduce the reference's print",
        not mismatches,
        "; ".join(mismatches) if mismatches
        else "three ratios over seven years, each rounding to the tenth IBCS "
             "printed",
    )

    # The tree itself: ROI is the product of the two middle boxes.
    worst = max(abs(C11A_ROS[i] / 100 * C11A_TURNOVER[i] * 100 - C11A_ROI[i])
                for i in range(n))
    check(
        "return on sales x capital turnover = return on investment, every year",
        worst < 1e-9,
        f"largest disagreement {worst:.2e} percentage points - the two routes "
        f"to ROI are return/capital and (return/sales)x(sales/capital), and "
        f"they are the same quotient",
    )

    # The links declared in the spec evaluate to the tiers they name.
    values = {"return": C11A_RETURN, "net_sales": C11A_NET_SALES,
              "capital": C11A_CAPITAL, "ros": C11A_ROS,
              "turnover": C11A_TURNOVER, "roi": C11A_ROI}
    scale = {"ros": 100.0, "roi": 100.0, "turnover": 1.0}
    broken = []
    for link in C11A_TREE.links:
        k = scale[link.result]
        for i in range(n):
            got = link.apply(values[link.left][i], values[link.right][i])
            if link.operator == ":":
                got *= k
            if abs(got - values[link.result][i]) > 1e-9:
                broken.append(f"{link.result} {C11A_YEARS[i]}")
    check(
        "every declared link computes the box it claims to produce",
        not broken,
        "three links over seven years" if not broken else ", ".join(broken),
    )

    # Scale groups. The certification requirement for this template is that
    # charts sharing a unit share a scale, so the groups have to be the ones
    # the units imply - not a drawing decision taken later.
    groups = C11A_TREE.groups()
    check(
        "the scale groups are the unit groups",
        groups == {"percent": ["roi", "ros"],
                   "turnover": ["turnover"],
                   "kEUR": ["return", "net_sales", "capital"]},
        "two percentage boxes, three currency boxes, and the dimensionless "
        "ratio alone - " + ", ".join(f"{k}:{len(v)}" for k, v in groups.items()),
    )

    # The message names the wrong year. Asserted, not corrected.
    oval = C11A.annotations[0].target[1]
    check(
        "the message's 19,2% is 2026, not the 2024 it names",
        round(C11A_ROI[5], 1) == 19.2 and round(C11A_ROI[3], 1) == 14.7
        and oval == 5,
        f"2026 ROI is {C11A_ROI[5]:.1f} and 2024's is {C11A_ROI[3]:.1f}; the "
        f"highlight oval sits on {C11A_YEARS[oval]}. The sentence is printed "
        f"as IBCS wrote it and this check records the discrepancy",
    )

    # The shared scale, checked against the drawing rather than asserted.
    # A group whose boxes really are one ruler is reproduced by one number; if
    # any box had been fitted to its own height instead, its bars would need a
    # different px-per-unit and the residual would blow out. Return peaks at
    # 5.5 where net sales reaches 27.7, so a per-box fit would have been the
    # easy mistake and this is what rules it out.
    values = {"return": C11A_RETURN, "net_sales": C11A_NET_SALES,
              "capital": C11A_CAPITAL, "ros": C11A_ROS,
              "turnover": C11A_TURNOVER, "roi": C11A_ROI}
    worst_key, worst_year, worst = "", "", 0.0
    counted = 0
    for group, keys in C11A_TREE.groups().items():
        k = C11A_SCALE_PX[group]
        for key in keys:
            for i, year in enumerate(C11A_YEARS):
                residual = C11A_MEASURED_EXTENTS[key][i] - values[key][i] * k
                counted += 1
                if abs(residual) > abs(worst):
                    worst_key, worst_year, worst = key, year, residual
    check(
        "one scale per unit reproduces all 42 measured bar extents within 1px",
        counted == 42 and abs(worst) < 1.0,
        f"{counted} bars off three rulers - {C11A_SCALE_PX['percent']:.4f} "
        f"px/%, {C11A_SCALE_PX['kEUR']:.4f} px/kEUR, "
        f"{C11A_SCALE_PX['turnover']:.4f} px/turn - worst residual "
        f"{worst:+.2f}px at {worst_key} {worst_year}",
    )

    # The plan boundary, which every box draws.
    check(
        "five actual years then two planned, in every box",
        C11A_SCENARIOS.count("AC") == 5 and C11A_SCENARIOS[-2:] == ("PL", "PL"),
        "the AC/PL rule falls between 2025 and 2026 in all six charts",
    )


# --------------------------------------------------------------------------- #
# C13D - small multiples, Beta Corp. 25 locations of Central Europe 2017..2028
# --------------------------------------------------------------------------- #
#
# The largest transcription in the library: fifteen panels of twelve years. It
# is also the best checked, because IBCS published the same data four ways -
# 13A and 13B as absolutes, 13C as an absolute variance and 13D as a relative
# one - and any two of those determine the third.

C13D_YEARS = tuple(str(y) for y in range(2017, 2029))
C13D_SCENARIOS = ("AC",) * 10 + ("PL",) * 2

C13D_PANELS = (
    ("cologne", "Cologne",
     (79, 73, 69, 53, 50, 21, 48, 49, 38, 39, 41, 41)),
    ("frankfurt", "Frankfurt",
     (50, 48, 40, 34, 89, 78, 38, 30, 24, 25, 24, 26)),
    ("munich", "Munich",
     (25, 26, 26, 45, 58, 88, 77, 66, 55, 44, 48, 50)),
    ("berlin", "Berlin",
     (66, 34, 28, 65, 136, 168, 190, 207, 344, 301, 309, 311)),
    ("zurich", "Zurich",
     (67, 69, 56, 84, 5, 97, 34, 102, 95, 102, 105, 107)),
    ("basel", "Basel",
     (88, 78, 67, 56, 43, 41, 52, 66, 69, 79, 84, 86)),
    ("lausanne", "Lausanne",
     (98, 88, 73, 79, 123, 66, 73, 70, 66, 62, 66, 68)),
    ("st_gallen", "St. Gallen",
     (6, 4, 12, 16, 36, 18, 10, 7, 14, 11, 21, 19)),
    ("vienna", "Vienna",
     (90, 94, 99, 34, 88, 93, 88, 65, 51, 44, 47, 49)),
    ("salzburg", "Salzburg",
     (7, 5, 25, 28, 56, 5, 9, 12, 14, 17, 19, 21)),
    ("linz", "Linz",
     (112, 104, 124, 113, 45, 123, 109, 98, 78, 89, 95, 97)),
    ("tail", "11 more small locations",
     (45, 49, 55, 52, 79, 55, 57, 47, 56, 54, 59, 61)),
    ("paris", "Paris",
     (78, 64, 98, 81, 71, 80, 59, 78, 69, 81, 86, 88)),
    ("lyon", "Lyon",
     (23, 34, 36, 46, 52, 69, 79, 84, 122, 89, 94, 96)),
    ("bordeaux", "Bordeaux",
     (68, 62, 69, 84, 60, 63, 72, 51, 70, 66, 65, 67)),
)

C13D_ABSOLUTE = {key: values for key, _label, values in C13D_PANELS}
C13D_LABELS = {key: label for key, label, _values in C13D_PANELS}

# What 13B prints in the tail panel it calls "12 more small locations".
# Transcribed separately, because it is the check on the two rosters rather
# than an input: St. Gallen plus the eleven-location panel has to equal it.
C13D_TAIL_12 = (51, 53, 67, 68, 115, 73, 67, 54, 70, 65, 80, 80)

# What the reference panel prints.
C13D_AVERAGE_PRINTED = (60, 55, 58, 58, 66, 71, 66, 69, 78, 74, 78, 79)


def _c13d_average() -> tuple[float, ...]:
    """The reference series: the mean of the fifteen panels.

    Not of twenty-five locations, though that is what the panel is titled. The
    divisor is the number of *panels*, one of which stands for eleven locations
    on its own - so eleven of the twenty-five are counted once between them.
    Dividing the same total by 25 reproduces none of the twelve printed
    figures; dividing by 15 reproduces all twelve. Asserted in the tie-outs
    rather than assumed here.
    """
    return tuple(sum(v[i] for v in C13D_ABSOLUTE.values()) / len(C13D_ABSOLUTE)
                 for i in range(len(C13D_YEARS)))


C13D_AVERAGE = _c13d_average()


def _c13d_variance(key: str) -> tuple[float, ...]:
    """One panel's relative variance from the reference, in percent.

    Computed as ``(v*15 - total) / total`` rather than as ``(v - total/15) /
    (total/15)``. The two are the same quantity, but the second divides twice
    and the error shows: Linz 2018 is 728/832 of the average, which is exactly
    87.5%, and going via the average lands 1.4e-14 short of it. That is enough
    to round the wrong way at a half boundary, and it is the only value in the
    template that sits on one.
    """
    count = len(C13D_ABSOLUTE)
    return tuple(
        (v * count - total) / total * 100.0
        for v, total in zip(
            C13D_ABSOLUTE[key],
            (sum(p[i] for p in C13D_ABSOLUTE.values())
             for i in range(len(C13D_YEARS)))))


C13D_VARIANCE = {key: _c13d_variance(key) for key in C13D_ABSOLUTE}

# The labels IBCS prints. All one hundred and eighty are exactly what the
# derivation produces - there is no exception, and the one that looked like an
# exception is worth keeping as a warning. Linz 2018 is 728/832 of the average,
# which is 87.5% exactly, and IBCS prints 88. Reached through the average it
# comes out 1.4e-14 short of the boundary and rounds to 87; rounded with
# Python's round() rather than Excel's rule it comes out 87 again, this time
# because 87.5 goes to even. Two independent ways to manufacture a discrepancy
# in the source that is not there.
C13D_PRINTED_VARIANCE = {
    "cologne": (31, 32, 18, -9, -24, -70, -28, -29, -51, -47, -47, -48),
    "frankfurt": (-17, -13, -32, -41, 35, 10, -43, -56, -69, -66, -69, -67),
    "munich": (-58, -53, -56, -22, -12, 24, 16, -4, -29, -40, -38, -37),
    "berlin": (10, -39, -52, 12, 106, 137, 186, 201, 343, 309, 299, 293),
    "zurich": (11, 24, -4, 45, -92, 37, -49, 48, 22, 39, 35, 35),
    "basel": (46, 41, 15, -3, -35, -42, -22, -4, -11, 7, 8, 9),
    "lausanne": (63, 59, 25, 36, 86, -7, 10, 2, -15, -16, -15, -14),
    "st_gallen": (-90, -93, -79, -72, -46, -75, -85, -90, -82, -85, -73, -76),
    "vienna": (50, 69, 69, -41, 33, 31, 33, -6, -34, -40, -39, -38),
    "salzburg": (-88, -91, -57, -52, -15, -93, -86, -83, -82, -77, -75, -73),
    "linz": (86, 88, 112, 95, -32, 73, 64, 42, 0, 21, 23, 23),
    "tail": (-25, -12, -6, -10, 20, -23, -14, -32, -28, -27, -24, -23),
    "paris": (30, 15, 68, 40, 7, 13, -11, 13, -11, 10, 11, 11),
    "lyon": (-62, -39, -38, -21, -21, -3, 19, 22, 57, 21, 21, 21),
    "bordeaux": (13, 12, 18, 45, -9, -11, 9, -26, -10, -10, -16, -15),
}


# Every pin stem, measured off the reference in pixels from its panel's
# zero line. None where the value is zero and no stem is drawn. Not an
# input to the drawing - this is how the one-scale claim is checked.
C13D_MEASURED_STEMS = {
    "cologne": (18, 18, 10, 5, 14, 41, 16, 17, 30, 27, 27, 28),
    "frankfurt": (9, 7, 18, 24, 20, 5, 25, 33, 40, 39, 40, 39),
    "munich": (34, 31, 32, 13, 7, 14, 9, 2, 17, 23, 22, 21),
    "berlin": (5, 22, 30, 7, 62, 80, 110, 119, 203, 183, 176, 173),
    "zurich": (6, 14, 2, 26, 54, 21, 28, 28, 13, 22, 20, 20),
    "basel": (27, 24, 8, 2, 20, 25, 12, 2, 6, 4, 4, 5),
    "lausanne": (37, 34, 14, 21, 51, 4, 5, 1, 8, 9, 8, 8),
    "vienna": (29, 41, 41, 24, 19, 18, 19, 3, 20, 23, 23, 22),
    "salzburg": (52, 53, 33, 30, 9, 55, 51, 48, 48, 45, 44, 43),
    "linz": (51, 51, 66, 56, 18, 43, 38, 25, None, 12, 13, 13),
    "tail": (14, 6, 3, 6, 11, 13, 8, 18, 16, 15, 14, 13),
    "paris": (17, 9, 40, 23, 4, 7, 6, 7, 6, 5, 6, 6),
    "lyon": (36, 22, 22, 12, 12, 1, 11, 13, 33, 12, 12, 12),
    "bordeaux": (7, 6, 10, 26, 5, 6, 5, 15, 5, 6, 9, 9),
}

# Fitted across all 167 of them: stem length = k*|variance| + c, where the
# intercept is the half-head the stem stops short of.
C13D_STEM_FIT = (0.59268, -0.534)


def _c13d_tier(key: str) -> Tier:
    """One panel: its relative variance, split into an actual and a plan run."""
    values = C13D_VARIANCE[key]
    return Tier(
        key=key,
        label=C13D_LABELS[key],
        kind="variance_rel",
        reference="AVG",
        number_format="{:+,.0f}",
        printed=C13D_PRINTED_VARIANCE[key],
        series=(
            Series("AC", _split(values, "AC", C13D_SCENARIOS)),
            Series("PL", _split(values, "PL", C13D_SCENARIOS)),
        ),
    )


C13D_REFERENCE_TIER = Tier(
    key="average",
    label="average 25 locations",
    kind="measure",
    number_format="{:,.0f}",
    printed=C13D_AVERAGE_PRINTED,
    series=(
        Series("AC", _split(C13D_AVERAGE, "AC", C13D_SCENARIOS)),
        Series("PL", _split(C13D_AVERAGE, "PL", C13D_SCENARIOS)),
    ),
)

# The reference's own arrangement: fourteen panels of pins with Berlin twice
# the height, and the reference panel drawn as columns on a shaded ground.
# St. Gallen is not drawn - it is inside the average, and the tail panel is
# relabelled to absorb it, which is the discrepancy the tie-outs record.
C13D_GRID_REFERENCE = PanelGrid(
    rows=4, cols=4, reference="average",
    cells=(
        PanelCell("cologne", 0, 0), PanelCell("frankfurt", 0, 1),
        PanelCell("munich", 0, 2), PanelCell("berlin", 0, 3, row_span=2),
        PanelCell("zurich", 1, 0), PanelCell("basel", 1, 1),
        PanelCell("lausanne", 1, 2),
        PanelCell("vienna", 2, 0), PanelCell("salzburg", 2, 1),
        PanelCell("linz", 2, 2),
        PanelCell("tail", 2, 3, label="12 more small locations"),
        PanelCell("paris", 3, 0), PanelCell("lyon", 3, 1),
        PanelCell("bordeaux", 3, 2),
        PanelCell("average", 3, 3, kind="column", shaded=True),
    ),
)

# The workbook's arrangement: one uniform 4x4 of pin panels, every cell the
# same size and the same kind. Berlin keeps its place in the grid rather than
# taking two rows, St. Gallen is drawn in its own right, and the tail panel
# reverts to the eleven locations it actually carries. The reference panel is
# not drawn at all - a pin is notation for a variance, and the baseline's
# variance from itself is a row of zeros - so it is stated in the subtitle and
# lives in the input block as the divisor. The sixteenth cell is left empty,
# which is where the reference puts Berlin's second row.
C13D_GRID_UNIFORM = PanelGrid(
    rows=4, cols=4, reference="average",
    cells=(
        PanelCell("cologne", 0, 0), PanelCell("frankfurt", 0, 1),
        PanelCell("munich", 0, 2), PanelCell("berlin", 0, 3),
        PanelCell("zurich", 1, 0), PanelCell("basel", 1, 1),
        PanelCell("lausanne", 1, 2), PanelCell("st_gallen", 1, 3),
        PanelCell("vienna", 2, 0), PanelCell("salzburg", 2, 1),
        PanelCell("linz", 2, 2), PanelCell("tail", 2, 3),
        PanelCell("paris", 3, 0), PanelCell("lyon", 3, 1),
        PanelCell("bordeaux", 3, 2), PanelCell(None, 3, 3),
    ),
)


C13D = Template(
    id="C13",
    variant="D",
    kind="small multiples",
    title=TitleBlock(
        entity="Beta Corp., 25 locations of Central Europe",
        measure="Net profit",
        unit="mEUR, relative variance in % from location average",
        period="2017..2028",
        message=("Berlin will further be above overall average in 2027 and 2028, "
                 "in 2028 its profits will be 293% above location average"),
    ),
    categories=C13D_YEARS,
    category_scenarios=C13D_SCENARIOS,
    tiers=tuple(_c13d_tier(key) for key in C13D_ABSOLUTE) + (C13D_REFERENCE_TIER,),
    panel_grids={"reference": C13D_GRID_REFERENCE, "uniform": C13D_GRID_UNIFORM},
    annotations=(Annotation(kind="oval", target=("berlin", 11)),),
    source_ref="template-refs/C13_13D.png",
    notes=(
        "Fifteen series and one reference, drawn two ways: the SVG follows the "
        "reference's own arrangement, the workbook a uniform grid of pins.",
        "Every panel is on one scale. That is the whole argument of a small "
        "multiple - a variance read in one panel has to mean the same in the "
        "next - and it is what Berlin's +343% costs the other fourteen.",
        "The panel titled 'average 25 locations' is the mean of the fifteen "
        "panels, not of twenty-five locations. Dividing by 25 reproduces none "
        "of its twelve printed figures; dividing by 15 reproduces all twelve.",
    ),
)


def _check_c13d(check) -> None:
    """Tie-outs for C13D.

    IBCS published this data four ways, so almost nothing here has to be taken
    on trust. The absolutes are transcribed; the reference series, every
    variance and the second roster are all derived from them, and each is
    checked against a figure IBCS printed somewhere else.
    """
    n = len(C13D_YEARS)

    check(
        "fifteen panels of twelve years",
        len(C13D_PANELS) == 15
        and all(len(v) == n for v in C13D_ABSOLUTE.values()),
        f"{len(C13D_PANELS)} panels x {n} years = "
        f"{len(C13D_PANELS) * n} transcribed values, the largest single "
        f"transcription in the library",
    )

    # The two rosters. 13A splits the tail as St. Gallen plus eleven more;
    # 13B folds both into twelve. If that does not close, the transcription of
    # either panel is wrong - and it closes on every one of the twelve years,
    # which is what licenses drawing the same data two ways.
    mismatched = [C13D_YEARS[i] for i in range(n)
                  if C13D_ABSOLUTE["st_gallen"][i] + C13D_ABSOLUTE["tail"][i]
                  != C13D_TAIL_12[i]]
    check(
        "St. Gallen plus the eleven-location panel is 13B's twelve-location one",
        not mismatched,
        "all twelve years reconcile, so 13A and 13B are one dataset seen two "
        "ways" if not mismatched else f"fails at {', '.join(mismatched)}",
    )

    # What the reference panel is the average of. This is the finding, and it
    # is stated as a comparison rather than an assertion so the numbers are on
    # the page: 15 reproduces every printed figure, 25 reproduces none.
    by15 = sum(round(sum(v[i] for v in C13D_ABSOLUTE.values()) / 15)
               == C13D_AVERAGE_PRINTED[i] for i in range(n))
    by25 = sum(round(sum(v[i] for v in C13D_ABSOLUTE.values()) / 25)
               == C13D_AVERAGE_PRINTED[i] for i in range(n))
    check(
        "'average 25 locations' is the mean of the fifteen panels, not of 25",
        by15 == n and by25 == 0,
        f"dividing the panel total by 15 reproduces {by15} of {n} printed "
        f"figures, by 25 reproduces {by25} - the divisor is the number of "
        f"panels, and the tail panel stands for eleven locations on its own",
    )

    # Every drawn pin against every printed label. Rounded Excel's way, which
    # matters exactly once and decisively: Linz 2018 is 728/832 of the average,
    # which is 87.5% *exactly*. Python's round() goes to even and lands on 87;
    # Excel rounds halves away from zero and lands on 88, which is what IBCS
    # printed. Round it the wrong way and this check reports a discrepancy in
    # the source that is really a discrepancy in the checker.
    off = []
    for key, printed in C13D_PRINTED_VARIANCE.items():
        for i in range(n):
            got = int(excel_round(C13D_VARIANCE[key][i]))
            if got != printed[i]:
                off.append(f"{C13D_LABELS[key]} {C13D_YEARS[i]}: derived "
                           f"{C13D_VARIANCE[key][i]:+.4f} prints {got:+d}, "
                           f"reference prints {printed[i]:+d}")
    check(
        "all 180 printed variances come back from the derivation",
        not off,
        f"{len(C13D_PRINTED_VARIANCE) * n} of "
        f"{len(C13D_PRINTED_VARIANCE) * n} exact, rounding halves away from "
        f"zero as Excel does. Linz 2018 is the one that turns on it: 728/832 "
        f"of the average is 87.5% exactly, which Excel prints 88 and Python's "
        f"round() prints 87" if not off else "; ".join(off),
    )

    # One scale across all fourteen pin panels, checked against the drawing.
    # Regressing 167 measured stem lengths on the derived variances gives a
    # single number with no residual over half a pixel - which is what says the
    # panels really are comparable, and what a per-panel fit would have broken.
    # (167 rather than 168 because Linz 2025 is exactly zero and draws no stem.)
    stems = C13D_MEASURED_STEMS
    k, c = C13D_STEM_FIT
    worst_key, worst_year, worst = "", "", 0.0
    for key, lengths in stems.items():
        for i, length in enumerate(lengths):
            if length is None:
                continue
            residual = length - (abs(C13D_VARIANCE[key][i]) * k + c)
            if abs(residual) > abs(worst):
                worst_key, worst_year, worst = key, C13D_YEARS[i], residual
    counted = sum(1 for v in stems.values() for x in v if x is not None)
    check(
        "one scale reproduces all 167 measured pin stems within half a pixel",
        counted == 167 and abs(worst) < 0.6,
        f"{counted} stems off a single {k:.5f}px per percentage point, worst "
        f"residual {worst:+.2f}px at {worst_key} {worst_year}. Berlin's +343% "
        f"is drawn on the same ruler as Salzburg's -15%, which is why it needs "
        f"two rows",
    )

    # The message names a figure the chart has to be able to show.
    berlin = C13D_VARIANCE["berlin"][-1]
    oval = C13D.annotations[0].target
    check(
        "the message's 293% is Berlin's last year, and the oval is on it",
        round(berlin) == 293 and oval == ("berlin", n - 1),
        f"Berlin 2028 is {berlin:.1f}% above the location average and the "
        f"highlight sits on {C13D_YEARS[oval[1]]}",
    )

    # Both grids describe the same sixteen cells, and between them account for
    # every panel exactly once in the workbook's case.
    uniform = C13D_GRID_UNIFORM
    drawn = {c.key for c in uniform.drawn()}
    check(
        "the uniform grid draws all fifteen panels and leaves one cell empty",
        drawn == set(C13D_ABSOLUTE) and len(uniform.cells) == 16
        and sum(1 for c in uniform.cells if c.key is None) == 1,
        f"{len(drawn)} panels in a {uniform.rows}x{uniform.cols} grid; the "
        f"reference panel is not among them because a pin is notation for a "
        f"variance and the baseline's variance from itself is zero",
    )

    reference = C13D_GRID_REFERENCE
    spanned = sum(c.row_span for c in reference.cells)
    check(
        "the reference grid gives Berlin two rows and draws the average as columns",
        reference.cell("berlin").row_span == 2
        and reference.cell("average").kind == "column"
        and reference.cell("average").shaded
        and spanned == 16,
        f"fifteen cells spanning {spanned} row-slots; Berlin reaches +343% "
        f"where no other panel passes +112%, and the reference panel is the "
        f"one thing on the page that is not a variance",
    )

    # The tail panel is relabelled in the reference arrangement, and that is a
    # discrepancy in the source rather than a choice of ours.
    check(
        "13D calls the tail panel '12 more' while drawing the eleven-location series",
        reference.cell("tail").label == "12 more small locations"
        and C13D_LABELS["tail"] == "11 more small locations",
        "13B's tail panel carries St. Gallen plus eleven and prints 51 in "
        "2017; 13C and 13D print -25%, which is the eleven-location panel's "
        "45 against the average of 60. Same title, different series - printed "
        "as IBCS printed it",
    )


# --------------------------------------------------------------------------- #
# Simple variants
# --------------------------------------------------------------------------- #
#
# IBCS's own lettered variants do not form a complexity gradient - they vary by
# scenario pair, sort order, measure basis and precision - so *simple* is a
# dimension this project defines rather than one it selects: the same chart and
# the same data, reduced to its base tier, with no variance tiers, no callouts
# and no highlight markers.
#
# For the tier templates that reduction is a matter of drawing fewer tiers, and
# it is expressed as a smaller SheetLayout in ibcs_layout - the Template keeps
# every tier, because `reference()`, the {ref} headers and the tie-outs all
# read them. Only the structure templates need the *data* reduced, because
# their panels are not tiers, and that is what this does.

SIMPLE_RULES = {
    # One panel, and only the measured years. C01A's argument is that three
    # panels share a scale; with one panel there is nothing to share it with,
    # so what is left is an ordinary stacked column chart - which is exactly
    # what someone reaching for "simple" wants.
    "C01A": {"panels": ("area",), "keep_scenarios": ("AC",)},
    # C02A is already a single panel. What makes it complex is the two subtotal
    # rows carrying the integrated legend, and the five callouts. Drop the
    # subtotals and it becomes a plain stacked bar - the horizontal counterpart
    # to C01A's simple form, which is a distinction IBCS treats as semantic:
    # time runs across, structure runs down.
    "C02A": {"drop_subtotals": True},
}


def _sliced(panel: "StructurePanel", keep: list[int]) -> "StructurePanel":
    return replace(
        panel,
        categories=tuple(panel.categories[i] for i in keep),
        category_scenarios=tuple(panel.category_scenarios[i] for i in keep),
        segments=tuple(replace(s, values=tuple(s.values[i] for i in keep))
                       for s in panel.segments),
        printed_totals=(None if panel.printed_totals is None
                        else tuple(panel.printed_totals[i] for i in keep)),
    )


def simplify(t: Template) -> Template:
    """The base-tier version of one template.

    Every template loses its callouts and comments, because a highlight is an
    argument about one number and the simple version is not making an argument.
    Beyond that, only the structure templates change: their panels are data,
    not layout, so a smaller chart means smaller data.
    """
    rule = SIMPLE_RULES.get(f"{t.id}{t.variant}", {})
    out = replace(t, annotations=(), comments=())
    if not rule:
        return out

    panels = t.structure_panels
    if rule.get("panels"):
        panels = tuple(p for p in panels if p.key in rule["panels"])

    rows = t.rows
    if rule.get("drop_subtotals") and rows:
        keep = [i for i, r in enumerate(rows) if r.kind != "subtotal"]
        rows = tuple(rows[i] for i in keep)
        panels = tuple(_sliced(p, keep) for p in panels)

    if rule.get("keep_scenarios"):
        wanted = set(rule["keep_scenarios"])
        panels = tuple(
            _sliced(p, [i for i, s in enumerate(p.category_scenarios)
                        if s in wanted])
            for p in panels)

    return replace(out, structure_panels=panels, rows=rows)


TEMPLATES: dict[str, Template] = {"C03A": C03A, "C04A": C04A,
                                 "C05X": C05X, "C06F": C06F,
                                 "C12A": C12A, "T01B": T01B,
                                 "T02A": T02A, "T03A": T03A,
                                 "T04A": T04A, "C01A": C01A, "C02A": C02A, "C07C": C07C, "C08H": C08H,
                                 "C09C": C09C, "C10D": C10D,
                                 "C11A": C11A,
                                 "C13D": C13D}

# Every template whose tie-outs live in their own function. C03 predates the
# convention and is still checked inline below.
_CHECKS = {"C04": _check_c04a, "C05": _check_c05x, "C06": _check_c06f,
           "C12": _check_c12a, "T01": _check_t01b,
           "T02": _check_t02a, "T03": _check_t03a, "T04": _check_t04a, "C01": _check_c01a, "C02": _check_c02a, "C07": _check_c07c, "C08": _check_c08h,
           "C09": _check_c09c, "C10": _check_c10d,
           "C11": _check_c11a,
           "C13": _check_c13d}


if __name__ == "__main__":
    # Tier labels carry Δ, and a failing check prints one. Windows consoles
    # default to cp1252, where that is fatal - so the report would die on the
    # one run where it has something to say.
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    for name, template in TEMPLATES.items():
        print(f"--- {name} " + "-" * (60 - len(name)))
        for result in check_ties(template, raise_on_fail=False):
            print(result)
        print()
    print("month   AC    PY(derived)   dPY(derived)  dPY printed   dPY%")
    for i, m in enumerate(MONTHS):
        print(
            f"{m:<6}{C03A_MEASURE[i]:>4}   {C03A_PY[i]:>10.1f}   "
            f"{C03A_VAR_ABS[i]:>11.1f}  {C03A_VAR_ABS_PRINTED[i]:>11d}   "
            f"{C03A_VAR_REL[i]:>6.1f}"
        )
