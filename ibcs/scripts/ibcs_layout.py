"""Per-template sheet layout: where the data zone ends and the chart zone begins.

Both things this module holds used to be C03A-shaped module constants inside
``ibcs_excel``, which meant template number two could not be built without
editing the renderer. They are per-template facts, so they live per template.

**The zone rule.** A sheet has exactly two reserved zones: a data zone of fixed
width on the left, and a chart zone to its right. Nothing may straddle them.
This is not cosmetic. The first version of the C03A workbook called
``Columns("A:N").AutoFit()`` with the IBCS message sentence sitting in F1, so
AutoFit sized column F to hold 140 characters - 544pt, against 250pt for columns
A to E combined - and the charts, pinned at a hard-coded 300pt, landed inside it.
That buried the message, the entire dPY% column feeding the top tier, and pushed
the chart-feed columns off to the right of the charts.

Nothing in the code had said where the data ended, so the charts had no way to
be placed after it. Now the widths are declared here, the chart zone starts
where the data zone stops, and ``verify()`` in the renderer asserts it.

Two consequences worth stating:

* **No AutoFit, anywhere.** A single long string anywhere in the block would
  silently move the zone boundary again.
* **The title block is merged.** Text in an unmerged cell spills across its
  empty neighbours, and the spill would run under the charts - hidden, because
  Excel draws shapes over cell text. Merging bounds it to the data zone. Merged
  cells do not auto-fit their row height, so the heights are set explicitly.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field, replace

import ibcs_data as D
import ibcs_style as S


# --------------------------------------------------------------------------- #
# The text block
# --------------------------------------------------------------------------- #
#
# Every word the chart shows comes from a cell. Someone who has only the .xlsx
# must be able to retype the unit once and have the subject line and every tier
# caption follow, without opening any source code - the same requirement that
# forced the numbers onto formulas, applied to the text.
#
# So the unit is its own input rather than a fragment of a sentence. The
# previous version wrote row 2 as "Contribution in kEUR", one string built in
# Python, which is editable but not *composable*: nothing else could refer to
# the unit alone, so a caption wanting "dPY kEUR" had to hard-code it.
#
# The reference scenario is an input for the same reason. It reads as structure
# rather than data, but it appears in the text of every variance caption, and
# changing PL to BU should not mean editing three text boxes.
TITLE_INPUTS = ("entity", "measure", "unit", "period", "reference", "message")

# Which of those the reader sees over the chart, in order. The subject line is
# derived rather than typed, so it is not an input.
TITLE_DISPLAY = ("entity", "subject", "period", "message")

# Caption patterns. ``{ref}`` and ``{unit}`` resolve to cell *references*, not to
# values - see SheetLayout.caption_formula. A tier with no caption gets no cell.
CAPTION_ABS = "Δ{ref} {unit}"
CAPTION_REL = "Δ{ref}%"

_PLACEHOLDER = re.compile(r"(\{ref\}|\{unit\})")


def col_letter(index: int) -> str:
    """1 -> A. Excel formulas are written with letters; the plan uses numbers."""
    letters = ""
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


@dataclass(frozen=True)
class Column:
    """One column of the data zone.

    ``header`` may contain ``{ref}``, which resolves to the scenario the variance
    tiers are measured against - PY for C03A, PL for C04A. Deriving it means a
    new template supplies data rather than a hand-written header list, and that
    the header cannot disagree with the tier it describes.

    ``width`` is in Excel's character units and is never computed from content.
    """

    key: str
    header: str
    width: float
    typed: bool = False          # shaded as an input; everything else is formula
    # Carried per column rather than applied to a range, because the two
    # templates do not order their columns the same way and a range like
    # "ref through var_abs" means different things in each.
    number_format: str | None = "0.0"
    # Engine columns the reader never edits and never needs to see. Hidden
    # rather than parked far right, and for a reason that turns out to matter:
    # `data_zone_width` measures the range the charts are positioned from, and
    # a hidden column measures zero - so the scale machinery can be added to a
    # sheet without moving a single chart on it.
    hidden: bool = False
    # The column this one is a scaled copy of, and the scale group whose span
    # it divides by. A chart series reads the scaled copy; nothing else does.
    scaled_from: str | None = None
    scale_group: str | None = None
    # The column this one is the printable text of. A data label linked to a
    # cell ignores the label's own number format and prints the cell at full
    # precision - proved in references/calibration - so the formatting has to
    # happen in the cell, which means TEXT() and a column to hold it.
    text_of: str | None = None
    text_format: str | None = None
    # A single formula in the first data row rather than a column of them.
    # The span cells are the only cells on the sheet that describe the whole
    # sheet rather than one period.
    single: bool = False
    # The column this one is a *clipped* copy of, and where to clip it. Used by
    # a tier that marks outliers instead of scaling to them - see
    # `clip_columns` for why those two jobs cannot be combined.
    clip_of: str | None = None
    clip_bounds: tuple[float, float] | None = None


@dataclass(frozen=True)
class TierSpec:
    """One tier chart: where it sits and what value range it spans.

    The field names are deliberately orientation-neutral, because a template
    stacks its tiers one way or the other and the same numbers mean different
    Excel properties in each case:

    ==============  ===================  =====================
    field           vertical (columns)   horizontal (bars)
    ==============  ===================  =====================
    ``offset``      chart Top            chart Left
    ``extent``      chart Height         chart Width
    ``plot_extent`` PlotArea.InsideHeight  PlotArea.InsideWidth
    ``plot_lead``   PlotArea.InsideTop   PlotArea.InsideLeft
    ==============  ===================  =====================

    ``cross_offset`` shifts the plot along the *shared category* axis, which is
    the other one in both cases.

    ``bounds`` is deliberately not fitted to the data. Tiers sharing a unit must
    share a scale, so the measure tier fixes units-per-point and every other tier
    in that unit derives its range from its own extent. ``check_scale`` proves it.
    """

    key: str
    offset: float
    extent: float
    bounds: tuple[float, float]
    plot_extent: float
    cross_offset: float = 0.0
    plot_lead: float = 6.0
    # How fat the elements are, in Excel's own terms. Per tier because a pin is
    # deliberately much thinner than a measure bar, and per template because the
    # category pitch differs - C03A has twelve categories where C04A has
    # nineteen, so the same GapWidth draws a different bar.
    gap_width: int | None = None
    overlap: int | None = None
    # The caption drawn over this tier, as a pattern in {ref} and {unit}. Held
    # here rather than taken from ``Tier.label`` because the Excel caption is a
    # formula over cells and the label is a finished string; ``check_captions``
    # proves the two still say the same thing. None means the tier gets no
    # caption - a measure tier is described by the subject line above it.
    caption: str | None = None
    # Which span this tier's values were divided by, where the sheet has a
    # scale block. The bounds below stay in data units - they are what the
    # geometry was measured in, and check_scale and the captions both read
    # them - and are divided by the span at the moment they reach the axis.
    scale_group: str | None = None
    # Stacked rather than clustered. A floating bar needs an invisible segment
    # under it, and a panel that has to draw five different fills on one axis
    # needs five series in one slot; both are stacks. A waterfall tier implies
    # it, so this is only set where the tier is not one.
    stacked: bool = False


@dataclass(frozen=True)
class SheetLayout:
    """Everything about one template's sheet that is not the data itself.

    ``orientation`` decides which axis the tiers stack along and which one they
    share. Vertical templates stack tiers down the page and share a category
    axis across it; horizontal templates put their tiers side by side and share
    a category axis down it. Every geometry field below is named for the shared
    axis rather than for x or y, so both cases read the same way.
    """

    template_id: str
    columns: tuple[Column, ...]
    formulas: dict[str, str]
    tiers: tuple[TierSpec, ...]
    n_categories: int
    orientation: str = "vertical"
    chart_span: float = 520.0        # the size every tier shares: width, or height
    chart_gap: float = 18.0          # points between the data zone and the charts
    chart_top: float = 10.0
    title_block_height: float = 76.0  # the title text boxes above the top tier
    plot_inside_lead: float = 46.0   # inset of the plot along the category axis
    plot_inside_span: float = 452.0  # size of the plot along the category axis
    same_unit_tiers: tuple[str, ...] = ("measure", "var_abs")
    # Which tier fixes units-per-point. Named rather than assumed: C12A's
    # base is a waterfall, and a template with no tier called "measure" was
    # silently skipping the shared-scale check when this was hard-coded.
    scale_base: str = "measure"
    # Whether the sheet carries the template's summary rows on its category
    # axis. Off by default: C03A and C04A draw their totals in the SVG only, and
    # switching them on would change two verified workbooks for no reason.
    include_summary: bool = False
    # Which tier holds the values behind a typed column, where the two are not
    # named the same. C05X types a column of variances called var_abs while its
    # variance tier is the waterfall - the renderer used to assume a column and
    # a tier with the same name, and a template that names them differently
    # could not be built at all.
    typed_source: dict = field(default_factory=dict)

    # ----------------------------------------------------------------- zones --
    @property
    def col(self) -> dict[str, int]:
        """key -> 1-based column number, in declaration order."""
        return {c.key: i for i, c in enumerate(self.columns, start=1)}

    @property
    def letters(self) -> dict[str, str]:
        return {k: col_letter(n) for k, n in self.col.items()}

    @property
    def last_column(self) -> int:
        return len(self.columns)

    @property
    def first_chart_column(self) -> int:
        """The first column the chart zone may occupy. verify() asserts this."""
        return self.last_column + 1

    # ------------------------------------------------------------- the rows --
    # The row plan is computed, not declared. It used to be ``first_row = 6``
    # with a comment saying rows 1-4 were the title, which stopped being true
    # the moment the title block gained a row - and a stale constant here moves
    # the whole data block under the headers silently.
    @property
    def title_rows(self) -> int:
        """The typed inputs, then the one derived line built from them."""
        return len(TITLE_INPUTS) + 1

    def title_row(self, field: str) -> int:
        """1-based row of one title input. ``subject`` is the derived line."""
        if field == "subject":
            return self.title_rows
        return TITLE_INPUTS.index(field) + 1

    @property
    def captioned(self) -> tuple[TierSpec, ...]:
        return tuple(t for t in self.tiers if t.caption)

    def caption_row(self, spec: TierSpec) -> int:
        # A blank row after the title block, so the two are read as two things.
        return self.title_rows + 2 + self.captioned.index(spec)

    def caption_formula(self, spec: TierSpec) -> str:
        '''Turn a caption pattern into an Excel formula over the title cells.

        ``"Δ{ref} {unit}"`` becomes ``="Δ"&$B$5&" "&$B$3``. Literal text is
        quoted, the placeholders become absolute references, and the result is
        what the caption *cell* holds - the text box on the chart then links to
        that cell, so there are two hops and only one source of truth.
        '''
        parts = []
        for token in _PLACEHOLDER.split(spec.caption):
            if not token:
                continue
            if token == "{ref}":
                parts.append(f"$B${self.title_row('reference')}")
            elif token == "{unit}":
                parts.append(f"$B${self.title_row('unit')}")
            else:
                parts.append('"' + token.replace('"', '""') + '"')
        return "=" + "&".join(parts)

    def caption_text(self, spec: TierSpec, template: D.Template) -> str:
        """The caption as it will read, for checking against the tier label."""
        return spec.caption.format(ref=self.reference(template),
                                   unit=template.title.unit)

    def header_row(self) -> int:
        # A blank row after the captions, for the same reason.
        return self.title_rows + 2 + len(self.captioned) + 1

    @property
    def first_row(self) -> int:
        return self.header_row() + 1

    def last_row(self) -> int:
        return self.first_row + self.n_categories - 1

    def headers(self, template: D.Template) -> list[str]:
        ref = self.reference(template)
        return [c.header.format(ref=ref) for c in self.columns]

    @staticmethod
    def reference(template: D.Template) -> str:
        """The scenario the variance tiers compare against - PY, PL or BU."""
        for tier in template.tiers:
            if tier.reference:
                return tier.reference
        raise ValueError(f"{template.id}{template.variant} has no reference scenario")

    def sheet_entries(self, template: D.Template) -> list[tuple[str, int]]:
        """The rows this sheet writes, as ("category" | "summary", index)."""
        if self.include_summary:
            return template.sheet_rows()
        return [("category", i) for i in range(len(template.categories))]

    def row_of(self, template: D.Template, kind: str, index: int) -> int:
        """The sheet row one entry lands on. Needed because C06F's variance
        rows are defined against totals rather than against a position."""
        for n, entry in enumerate(self.sheet_entries(template)):
            if entry == (kind, index):
                return self.first_row + n
        raise KeyError(f"{kind} {index} is not on this sheet")

    def structure_values(self, template: D.Template, index: int,
                         kind: str = "category") -> dict:
        """The columns that describe a row rather than measure it.

        Which ones those are is a per-template fact: a time series has a period
        and a scenario, a statement has a line and a sign. Resolved from the
        column plan rather than assumed, because the renderer used to write
        ``col["period"]`` unconditionally and a template without one could not
        be built at all.
        """
        keys, out = set(self.col), {}
        if kind == "summary":
            row = template.summary_rows[index]
            if "line" in keys:
                out["line"] = row.label
            if "period" in keys:
                out["period"] = row.label
            if "scenario" in keys:
                out["scenario"] = row.stack[0][0] if row.stack else ""
            if "kind" in keys:
                out["kind"] = "variance" if row.is_variance_only else "total"
            if "sign" in keys:
                out["sign"] = 1
            return out

        if "period" in keys:
            out["period"] = template.categories[index]
        if "scenario" in keys:
            out["scenario"] = template.category_scenarios[index]
        if "line" in keys:
            row = template.rows[index]
            out["line"] = f"{row.prefix} {row.label}" if row.prefix else row.label
        if "kind" in keys:
            out["kind"] = "state"
        if "sign" in keys:
            out["sign"] = template.rows[index].sign
        return out

    def typed_override(self, template: D.Template, index: int, key: str,
                       row: int, kind: str = "category") -> str | None:
        """A typed column is not typed on every row.

        On a statement the subtotals are consequences of the lines above them,
        not inputs. Typing them lets a component change without its total
        following, which is the one thing a statement must never do - and it is
        invisible, because the *bar* is drawn from the walk and stays right
        while the number printed beside it goes stale.

        Both forms reference only rows above, so nothing is circular: a result
        line reads the level the statement had reached, and a group line the
        distance the block covered.
        """
        if kind == "summary":
            return self._summary_override(template, index, key)
        if not template.rows or key not in ("py", "ac"):
            return None
        r = template.rows[index]
        if r.kind != "subtotal" or index == 0:
            return None
        level = self.letters[f"{key}_level"]
        if r.spans == "zero":
            return f"={level}{row - 1}"
        return f"=ABS({level}{row - 1}-{level}{row - int(r.spans) - 1})"

    def _summary_override(self, template: D.Template, index: int,
                          key: str) -> str | None:
        """A total that has components is a formula, not an input.

        The same rule C12A's subtotals proved, one template later and with the
        failure visible this time: while C06F's actual total was typed, raising
        a state left the total where it was and the bridge stopped landing on
        it. The chart said the walk ended somewhere the bar did not.

        The cost is two labels: the derived totals come to 1 730 and 2 074 where
        IBCS printed 1 728 and 2 071, because fifteen rounded states do not add
        to a rounded total. That is the documented split - the SVG prints what
        the original printed, the workbook prints what its own data says, and a
        workbook that disagrees with itself would be worse than one that
        disagrees with the reference by two.
        """
        if key != "measure" or "wf_base" not in self.col:
            return None
        srow = template.summary_rows[index]
        if not srow.stack:
            return None
        # Which column holds this scenario's monthly figures. Asking the column
        # plan rather than assuming: C06F derives its prior-year total from a
        # "py" column and C05X its plan total from a "pl" one, and each has a
        # scenario the other does not carry - C06F's plan and C05X's prior year
        # have no components on the sheet, so they stay typed.
        column_key = {"AC": "measure", "PY": "py", "PL": "pl"}.get(srow.stack[0][0])
        if column_key is None or column_key not in self.col:
            return None
        entries = self.sheet_entries(template)
        rows = [self.first_row + n for n, (k, _) in enumerate(entries)
                if k == "category"]
        column = self.letters[column_key]
        return f"=SUM({column}{rows[0]}:{column}{rows[-1]})"

    def row_formulas(self, template: D.Template, index: int, row: int,
                     kind: str = "category") -> dict[str, str]:
        """Formulas that depend on which row this is, not just which column.

        Empty for every template whose rows are interchangeable. A statement
        overrides it, because a subtotal's bar does not start where an ordinary
        line's does and no single template string covers both.
        """
        if "wf_up_ac" in self.col:
            return c05x_formulas(self, template, index, row, kind)
        if "wf_level" in self.col:
            return c06f_formulas(self, template, index, row, kind)
        if kind != "category" or not template.rows or "py_level" not in self.col:
            return {}
        return waterfall_formulas(self, template, index, row)

    def variance_splits(self, template: D.Template, key: str) -> list:
        """The series a variance tier is split into, and each one's colour.

        Two for a tier whose impact direction is fixed, four where it varies by
        row. Returned as (column key, colour) so the renderers do not have to
        know which case they are in - they iterate whatever they are given.
        """
        stem = "abs" if key == "var_abs" else "rel"
        if f"{stem}_pos_good" in self.col:
            return [(f"{stem}_pos_good", S.VARIANCE["good"]),
                    (f"{stem}_pos_bad", S.VARIANCE["bad"]),
                    (f"{stem}_neg_good", S.VARIANCE["good"]),
                    (f"{stem}_neg_bad", S.VARIANCE["bad"])]
        higher_is_better = template.tier(key).higher_is_better
        return [(f"{stem}_up", S.variance_colour(1, higher_is_better)),
                (f"{stem}_dn", S.variance_colour(-1, higher_is_better))]

    def formula(self, key: str, row: int) -> str:
        """Resolve one cell formula. Column letters come from the plan, never
        from a literal, so reordering the plan cannot silently break a formula."""
        return self.formulas[key].format(r=row, **self.letters)

    # ------------------------------------------------------------- scaling --
    @property
    def scale_groups(self) -> tuple[str, ...]:
        seen = []
        for c in self.columns:
            if c.single and c.scale_group and c.scale_group not in seen:
                seen.append(c.scale_group)
        return tuple(seen)

    def span_ref(self, group: str) -> str:
        """Absolute address of one group's span cell."""
        key = f"span_{group}"
        return f"${self.letters[key]}${self.first_row}"

    def span_formula(self, group: str) -> str:
        """The largest magnitude any chart in this group has to draw.

        Not rounded to anything tidy, and it does not need to be: the value
        axis is stripped from every one of these charts, so nobody ever reads a
        tick off it. The span exists to be divided by, not to be looked at.

        AGGREGATE rather than MAX because the feed columns are full of NA() -
        a slot that does not apply on this row - and MAX over an error is an
        error. Options 6 tells it to ignore them. Guarded away from zero so a
        sheet of blanks divides by something.
        """
        # The *sources*, not the scaled copies. Reading the scaled columns
        # would make the span depend on itself, and Excel would answer with a
        # circular-reference warning and a sheet full of zeros.
        cols = [self.letters[c.scaled_from] for c in self.columns
                if c.scaled_from and c.scale_group == group]
        span = ",".join(f"{c}{self.first_row}:{c}{self.last_row()}"
                        for c in cols)
        return (f"=MAX(AGGREGATE(4,6,{span}),-AGGREGATE(5,6,{span}),1E-9)")

    def scale_formula(self, column: "Column", row: int) -> str:
        """One scaled or text cell.

        The scaled copy divides; the text copy is what the data label reads,
        and it has to be text rather than a number because a linked label
        ignores its own number format.
        """
        src = self.letters[column.clip_of or column.scaled_from or column.text_of]
        if column.clip_of:
            # MEDIAN of three is the shortest clamp Excel has. NA() stays NA():
            # a slot that does not apply on this row must draw nothing, and
            # MEDIAN would otherwise turn it into a bar at the boundary.
            lo, hi = column.clip_bounds
            return (f"=IF(ISNA({src}{row}),NA(),"
                    f"MEDIAN({lo},{src}{row},{hi}))")
        if column.scaled_from:
            return f"={src}{row}/{self.span_ref(column.scale_group)}"
        # A number format may itself contain quotes - the waterfall's deduction
        # segments print a literal minus in front of a magnitude, which is
        # written "-"#,##0. Inside a formula string each of those has to be
        # doubled or the formula ends early and Excel rejects the whole cell.
        fmt = column.text_format.replace('"', '""')
        return f'=IF(ISNA({src}{row}),"",TEXT({src}{row},"{fmt}"))'

    def drawn_column(self, key: str) -> str:
        """The column a chart series should actually read for `key`.

        The scaled copy where one exists, the column itself where it does not -
        so a sheet that has not been given a scale block still builds.
        """
        if f"c_{key}" in self.col:
            return f"c_{key}"
        return f"s_{key}" if f"s_{key}" in self.col else key

    def label_column(self, key: str) -> str | None:
        """The column a data label should link to, if the values are scaled."""
        return f"t_{key}" if f"t_{key}" in self.col else None

    @property
    def horizontal(self) -> bool:
        return self.orientation == "horizontal"

    def tier_offset(self, tier: TierSpec) -> float:
        """Where a tier's chart object starts along the stacking axis.

        The title block sits above the charts either way, so a vertical
        template's tiers all shift down by it and a horizontal template's tiers
        all start below it - equally, so alignment is unaffected.
        """
        if self.horizontal:
            return tier.offset
        return self.chart_top + self.title_block_height + tier.offset

    # ------------------------------------------------------------- geometry --
    def category_offset(self, fraction: float = 9 / 56) -> float:
        """A clustered column sits inside its category slot; a line/marker series
        sits on the midpoint. The measure tier puts the reference scenario on the
        midpoint and the measure to its right, so anything aligning with the
        measure bars moves right by that fraction of a category - 9px of a 56px
        pitch in the reference render."""
        return self.plot_inside_span / self.n_categories * fraction

    def check_scale(self) -> list[str]:
        """Prove the same-unit tiers really are drawn at one scale.

        The measure tier sets the exchange rate; every other same-unit tier's
        range must equal its own plot height times that rate. Getting this wrong
        produces a chart that looks fine and exaggerates every variance on it,
        which is precisely the error the rule exists to prevent.
        """
        tiers = {t.key: t for t in self.tiers}
        base = tiers.get(self.scale_base)
        if base is None or self.scale_base not in self.same_unit_tiers:
            return []
        lo, hi = base.bounds
        rate = (hi - lo) / base.plot_extent
        problems = []
        for key in self.same_unit_tiers:
            if key == self.scale_base:
                continue
            tier = tiers[key]
            wanted = tier.plot_extent * rate
            actual = tier.bounds[1] - tier.bounds[0]
            if abs(actual - wanted) > 0.01:
                problems.append(
                    f"{self.template_id} tier {key}: spans {actual:g} units over "
                    f"{tier.plot_extent:g}pt, but the measure tier's scale of "
                    f"{rate:.4f} units/pt requires {wanted:g}")
        return problems

    def check_captions(self, template: D.Template) -> list[str]:
        """Prove the Excel caption and the SVG tier label still agree.

        Excel builds its caption from cells and the SVG prints ``Tier.label``, so
        the two engines could drift apart and each look right on its own. The
        unit is the part Excel adds, so it is stripped before comparing - what
        must match is the scenario notation, which is the part that carries
        meaning.
        """
        problems = []
        for spec in self.captioned:
            derived = self.caption_text(spec, template)
            without_unit = derived.replace(template.title.unit, "").strip()
            label = template.tier(spec.key).label
            if without_unit != label:
                problems.append(
                    f"{self.template_id} tier {spec.key}: the caption cell will "
                    f"read {derived!r} but the SVG prints {label!r} - the two "
                    f"renderers would disagree")
        return problems


# --------------------------------------------------------------------------- #
# The chart-feed formulas
# --------------------------------------------------------------------------- #
#
# Which two columns are typed is a per-template fact, not a constant. C03A
# transcribes the measure and the relative variance, so the reference scenario
# is derived; C04A transcribes the measure and the *absolute* variance, so the
# reference scenario is derived from that instead. Everything else in either set
# follows, so someone can retype a number and the whole chart follows without
# opening any source code.
#
# The up/down pairs exist because Excel cannot colour a point by its value - it
# colours a whole series. A variance that must be green when positive and red
# when negative is therefore plotted as two mutually exclusive series.
#
# NA() rather than "" or 0: a blank plots as a gap, an empty string can coerce to
# zero, and a zero draws a bar of no height sitting on the axis. NA() is the only
# one Excel treats as "this point does not exist".
#
# No rounding anywhere. Pre-rounding a variance to one decimal is what made
# November read -13 where IBCS printed -12: it compounds with Excel's own
# rounding, which takes halves away from zero. Full precision in the cell, and
# let the number format do the displaying.
SPLIT_FORMULAS = {
    "abs_up": "=IF(${var_abs}{r}>0,${var_abs}{r},NA())",
    "abs_dn": "=IF(${var_abs}{r}<0,${var_abs}{r},NA())",
    "rel_up": "=IF(${var_rel}{r}>0,${var_rel}{r},NA())",
    "rel_dn": "=IF(${var_rel}{r}<0,${var_rel}{r},NA())",
    "rel_up_len": "=IF(${var_rel}{r}>0,${var_rel}{r},0)",
    "rel_dn_len": "=IF(${var_rel}{r}<0,-${var_rel}{r},0)",
}

# On a statement the impact direction is a property of the line, not of the
# tier: within one dPY column, revenue up is favourable and cost up is not. A
# two-way split cannot say that, because Excel colours a series and the sign of
# the value decides which side of the axis the bar falls on - so the four
# combinations of (which side, which impact) need four series.
#
# The sign column is what makes it live. Retype a figure and the bar moves and
# recolours; change a line from a cost to a credit and it recolours without
# moving. Neither is possible with per-point formatting, which freezes at build.
IMPACT_SPLIT_FORMULAS = {
    "abs_pos_good": "=IF(AND(${var_abs}{r}>0,${sign}{r}>0),${var_abs}{r},NA())",
    "abs_pos_bad": "=IF(AND(${var_abs}{r}>0,${sign}{r}<0),${var_abs}{r},NA())",
    "abs_neg_good": "=IF(AND(${var_abs}{r}<0,${sign}{r}<0),${var_abs}{r},NA())",
    "abs_neg_bad": "=IF(AND(${var_abs}{r}<0,${sign}{r}>0),${var_abs}{r},NA())",
    "rel_pos_good": "=IF(AND(${var_rel}{r}>0,${sign}{r}>0),${var_rel}{r},NA())",
    "rel_pos_bad": "=IF(AND(${var_rel}{r}>0,${sign}{r}<0),${var_rel}{r},NA())",
    "rel_neg_good": "=IF(AND(${var_rel}{r}<0,${sign}{r}<0),${var_rel}{r},NA())",
    "rel_neg_bad": "=IF(AND(${var_rel}{r}<0,${sign}{r}>0),${var_rel}{r},NA())",
}

TIER_FORMULAS = {
    "var_abs": "={measure}{r}-{ref}{r}",
    "var_rel": "=IF({ref}{r}=0,NA(),({measure}{r}-{ref}{r})/{ref}{r}*100)",
    "meas_ac": '=IF(${scenario}{r}="AC",${measure}{r},NA())',
    "meas_fc": '=IF(${scenario}{r}="AC",NA(),${measure}{r})',
    "abs_up": "=IF(${var_abs}{r}>0,${var_abs}{r},NA())",
    "abs_dn": "=IF(${var_abs}{r}<0,${var_abs}{r},NA())",
    "rel_up": "=IF(${var_rel}{r}>0,${var_rel}{r},NA())",
    "rel_dn": "=IF(${var_rel}{r}<0,${var_rel}{r},NA())",
    "rel_up_len": "=IF(${var_rel}{r}>0,${var_rel}{r},0)",
    "rel_dn_len": "=IF(${var_rel}{r}<0,-${var_rel}{r},0)",
}

# The six columns a reader actually reads, then the chart-feed scaffolding. The
# scaffolding is narrow with a wrapped header rather than hidden: it is how the
# chart is fed, and someone tracing a bar back to a number should be able to.
SIGNED = "+0.0;-0.0"

def scale_columns(plan, groups) -> tuple["Column", ...]:
    """The hidden half of a data zone: one span per group, and per drawn
    column a scaled copy for the chart and a text copy for its label.

    Three columns where there was one, and each does a job the others cannot.
    Excel will not bind an axis bound to a formula, so a chart cannot be made
    to follow its data by moving the axis; the data has to be moved instead.
    Dividing by a span means the value plotted is a *fraction*, the axis is a
    fixed set of fractions, and any figures at all land inside it.

    Dividing rather than shifting and scaling is deliberate: zero divided by
    anything is still zero, so the zero line, the reference rule and every
    caption positioned from the plot geometry stay exactly where they were.

    The text copy exists because a data label linked to a cell **ignores the
    label's own number format** and prints the cell at full precision - proved
    in references/calibration/calib_label.py, and re-applying the format after
    linking does not help. So the formatting has to happen in the cell.
    """
    out = [Column(f"span_{g}", f"span {g}", 8.0, number_format="0.000000",
                  hidden=True, single=True, scale_group=g) for g in groups]
    for key, group, fmt in plan:
        out.append(Column(f"s_{key}", f"{key} scaled", 8.0,
                          number_format="0.000000", hidden=True,
                          scaled_from=key, scale_group=group))
        if fmt:
            # General, not "@". A text format applied before the formula is
            # written makes Excel store the formula as literal text, and the
            # data label then prints IF(ISNA(...)) across the chart. The cell
            # already returns a string; it does not need telling.
            out.append(Column(f"t_{key}", f"{key} text", 8.0,
                              number_format=None, hidden=True,
                              text_of=key, text_format=fmt))
    return tuple(out)


def clip_columns(plan, bounds) -> tuple["Column", ...]:
    """The other answer to "somebody else's data": clip it and say so.

    `scale_columns` makes a tier follow its data by dividing everything by a
    span. That works whenever a template's declared range is *wider* than its
    data - the usual case, where the range is headroom around the numbers.

    It cannot work when the declared range is deliberately **narrower** than the
    data, which is what an outlier-marking tier is. The axis is fixed at build,
    so a series plotting `v / span` puts the largest value at exactly 1.0 every
    time - it *is* the span - and a tier whose axis stops at a quarter of that
    clips its own span-defining value forever, on any data at all. C12A's
    relative tier stops at +265% over data reaching +983%; normalising it would
    make the biggest bar run off whatever the figures were. Following the data
    and clipping outliers are contradictory jobs and this tier's job is the
    second one.

    So: no span, no division. A clipped copy for the chart to draw, a text copy
    carrying the value that was actually measured, and - written by the caller -
    a marker column saying an element ran off and by how much. All three are
    formulas, so all three follow an edit.
    """
    lo, hi = bounds
    out = []
    for key, fmt in plan:
        out.append(Column(f"c_{key}", f"{key} clipped", 7.0,
                          number_format=None, hidden=True,
                          clip_of=key, clip_bounds=(lo, hi)))
        if fmt:
            out.append(Column(f"t_{key}", f"{key} text", 7.0,
                              number_format=None, hidden=True,
                              text_of=key, text_format=fmt))
    return tuple(out)


STANDARD_COLUMNS = (
    Column("period", "Period", 8.0, number_format=None),
    Column("scenario", "Scenario", 9.0, number_format=None),
    Column("ref", "{ref}", 8.5, typed=True),
    Column("measure", "Measure", 9.0, typed=True),
    Column("var_abs", "Δ{ref}", 8.5),
    Column("var_rel", "Δ{ref}%", 8.5, number_format=SIGNED),
    Column("meas_ac", "Measure AC", 7.5, number_format=SIGNED),
    Column("meas_fc", "Measure FC", 7.5, number_format=SIGNED),
    Column("abs_up", "Δ{ref} up", 7.5, number_format=SIGNED),
    Column("abs_dn", "Δ{ref} down", 7.5, number_format=SIGNED),
    Column("rel_up", "Δ{ref}% up", 7.5, number_format=SIGNED),
    Column("rel_dn", "Δ{ref}% down", 7.5, number_format=SIGNED),
    Column("rel_up_len", "up length", 7.5, number_format=SIGNED),
    Column("rel_dn_len", "down length", 7.5, number_format=SIGNED),
) + scale_columns(
    # (drawn column, its scale group, the format its label prints in)
    # Exactly the columns a chart series reads, and nothing else. The pin
    # stems share the percentage span with the pin heads because they are
    # lengths on the same axis; scaling them apart would leave every stem
    # pointing at the wrong place.
    (("ref", "unit", None),
     ("measure", "unit", "0"),
     ("var_abs", "unit", "+0;-0"),
     ("abs_up", "unit", None),
     ("abs_dn", "unit", None),
     ("rel_up", "rel", "+0.0;-0.0"),
     ("rel_dn", "rel", "+0.0;-0.0"),
     ("rel_up_len", "rel", None),
     ("rel_dn_len", "rel", None)),
    groups=("unit", "rel"))


# --------------------------------------------------------------------------- #
# C03A - multi-tier columns, twelve months
# --------------------------------------------------------------------------- #
#
# The value-axis bounds below are the reference render's, kept as literals. The
# range of the absolute-variance tier is exactly derivable - 84pt at the measure
# tier's 1.0952 units/pt is 92.0 units, and 55.5 - -36.5 is 92.0, which
# check_scale() asserts. Only the split of that range about zero is a rounding of
# the render's 41:27 proportion (55.47/-36.53), and it is kept as drawn.
#
# The relative tier is a percentage - a different unit - so it is free.

_C03A_OFFSET = 452.0 / 12 * (9 / 56)      # see SheetLayout.category_offset

C03A = SheetLayout(
    template_id="C03A",
    columns=STANDARD_COLUMNS,
    formulas=TIER_FORMULAS,
    n_categories=12,
    tiers=(
        TierSpec("var_rel", 10.0, 104.0, (-15.0, 23.0), 84.0, _C03A_OFFSET,
                 caption=CAPTION_REL, scale_group="rel"),
        TierSpec("var_abs", 118.0, 104.0, (-36.5, 55.5), 84.0, _C03A_OFFSET,
                 gap_width=47, overlap=100, caption=CAPTION_ABS,
                 scale_group="unit"),
        TierSpec("measure", 226.0, 244.0, (0.0, 230.0), 210.0, 0.0,
                 gap_width=23, overlap=76, scale_group="unit"),
    ),
)


# --------------------------------------------------------------------------- #
# C04A - multi-tier bars, nineteen states against plan
# --------------------------------------------------------------------------- #
#
# The transpose of C03A, and the test of whether these abstractions generalise.
# Three things it does that C03A does not:
#
#   * tiers side by side rather than stacked, sharing a *vertical* category axis;
#   * the reference scenario derived rather than typed - IBCS prints AC and the
#     absolute variance here, so PL = AC - dPL;
#   * a dPL reference axis, which is a double rule rather than a single line.
#
# Geometry is converted from the SVG at 0.75 pt/px, then rounded so the shared
# scale comes out exact: the measure tier spans 240 kUSD over 288pt, which is
# 0.8333 units/pt, and the dPL tier's 180pt therefore spans exactly 150.
# check_scale() asserts it rather than trusting the arithmetic above.

C04A_FORMULAS = {
    "ref": "={measure}{r}-{var_abs}{r}",
    "var_rel": "=IF({ref}{r}=0,NA(),{var_abs}{r}/{ref}{r}*100)",
    **SPLIT_FORMULAS,
}

# No forecast columns: every category in C04A is an actual, so the AC/FC split
# that C03A needs would be two columns of constants.
C04A_COLUMNS = (
    Column("period", "State", 15.0, number_format=None),
    Column("scenario", "Scenario", 9.0, number_format=None),
    Column("measure", "Measure", 9.0, typed=True, number_format="0"),
    Column("var_abs", "Δ{ref}", 8.5, typed=True, number_format="+0;-0"),
    Column("ref", "{ref}", 8.5, number_format="0"),
    Column("var_rel", "Δ{ref}%", 8.5, number_format=SIGNED),
    Column("abs_up", "Δ{ref} up", 7.5, number_format="+0;-0"),
    Column("abs_dn", "Δ{ref} down", 7.5, number_format="+0;-0"),
    Column("rel_up", "Δ{ref}% up", 7.5, number_format=SIGNED),
    Column("rel_dn", "Δ{ref}% down", 7.5, number_format=SIGNED),
    Column("rel_up_len", "up length", 7.5, number_format=SIGNED),
    Column("rel_dn_len", "down length", 7.5, number_format=SIGNED),
) + scale_columns(
    # Which column each series actually reads, and the format its label prints
    # in where it shows one. The horizontal pin differs from C03A's vertical
    # one: there the two split columns carry the markers and the labels, here
    # an invisible `var_rel` labeller carries them and the splits only colour
    # the stems. So the text copy belongs to `var_rel`, not to the splits.
    #
    # `rel_up_len`/`rel_dn_len` are absent deliberately. They feed the error
    # bars of the *vertical* pin and nothing on a horizontal sheet reads them;
    # scaling them would put their magnitudes into the span for no reason.
    (("ref", "unit", None),
     ("measure", "unit", "0"),
     ("var_abs", "unit", "+0;-0"),
     ("abs_up", "unit", None),
     ("abs_dn", "unit", None),
     ("var_rel", "rel", "+0;-0"),
     ("rel_up", "rel", None),
     ("rel_dn", "rel", None)),
    groups=("unit", "rel"))

C04A = SheetLayout(
    template_id="C04A",
    columns=C04A_COLUMNS,
    formulas=C04A_FORMULAS,
    n_categories=19,
    orientation="horizontal",
    chart_span=400.0,             # every panel is this tall
    plot_inside_lead=20.0,        # the category axis starts here, in every panel
    plot_inside_span=365.0,       # 19 rows x 25.611px at 0.75 pt/px
    tiers=(
        # key,      left,  width, bounds,          plot width, cross, plot lead
        TierSpec("measure", 0.0, 430.0, (0.0, 240.0), 288.0, 0.0, 110.0,
                 gap_width=20, overlap=76, scale_group="unit"),
        TierSpec("var_abs", 435.0, 250.0, (-105.0, 45.0), 180.0, 0.0, 40.0,
                 gap_width=35, overlap=100, caption=CAPTION_ABS,
                 scale_group="unit"),
        # The pin stem. GapWidth is what makes it thin, and 400 on a nineteen
        # category axis lands close to the reference's 5px stem.
        TierSpec("var_rel", 690.0, 170.0, (-65.0, 35.0), 75.0, 0.0, 55.0,
                 gap_width=500, overlap=100, caption=CAPTION_REL,
                 scale_group="rel"),
    ),
)


# --------------------------------------------------------------------------- #
# C12A - a profit and loss statement as two waterfalls
# --------------------------------------------------------------------------- #
#
# The first template whose panels are not all the same kind of chart: two
# waterfalls, then the two variance tiers between them. Four panels, three of
# them in kEUR and therefore on one scale, which check_scale still enforces -
# the base tier is named rather than assumed to be "measure".
#
# The scale is chosen so the shared-unit arithmetic comes out exact rather than
# nearly exact: 4 kEUR per point, so the waterfall panels' 1 120 kEUR span 280pt
# and the dPY panel's 130pt spans exactly 520.
#
# What the workbook does that the SVG does not have to: keep the walk live. The
# running level is a formula column, so retyping any line re-floats every bar
# below it. A column of static bases would look identical on the build data and
# start lying on the first edit - the same failure the variance split fixed for
# colour, one level further down.
#
# The structure - which line adds, which is a subtotal, and what a subtotal
# summarises - is written as constants, not formulas. That is deliberate and is
# the same limit C03A documents for its AC/FC split: values respond to an edit,
# structure does not, because a statement's shape is not data.

# Which preceding level a row's bar starts from. Elements start where the last
# line left off; a result subtotal starts at zero; a group subtotal starts where
# its block began, which is `spans` rows further back.
def _c12a_start(row, r: int, level: str, index: int) -> str:
    if row.kind != "subtotal":
        # The first line starts from nothing; every other one starts where the
        # previous left off. Without the index this reached a row above the
        # data block - the header - and read it as a level of zero by luck.
        return "=0" if index == 0 else f"={level}{r - 1}"
    if row.spans == "zero":
        return "=0"
    return f"={level}{r - int(row.spans) - 1}"


# The relative tier's declared range, and where a clipped bar stops inside it.
# Ninety per cent of each side, so the marker has somewhere to sit - the SVG
# reserves the same gap by backing off from its right-hand limit.
C12A_REL_LIMIT = (-35.0, 265.0)
C12A_REL_STOP = (C12A_REL_LIMIT[0] * 0.9, C12A_REL_LIMIT[1] * 0.9)

C12A_COLUMNS = (
    Column("line", "Line", 22.0, number_format=None),
    Column("sign", "Sign", 5.0, number_format="+0;-0"),
    Column("py", "PY", 8.0, typed=True, number_format="#,##0"),
    Column("ac", "AC", 8.0, typed=True, number_format="#,##0"),
    Column("var_abs", "Δ{ref}", 8.0, number_format="+#,##0;-#,##0"),
    Column("var_rel", "Δ{ref}%", 8.5, number_format=SIGNED),
    # The walk, one column per scenario. Visible rather than hidden: a reader
    # tracing a floating bar back to a number should be able to find the level
    # it floats on.
    Column("py_level", "PY level", 8.0, number_format="#,##0"),
    Column("py_from", "PY from", 8.0, number_format="#,##0"),
    Column("py_base", "PY base", 8.0, number_format="#,##0"),
    Column("py_add", "PY adds", 8.0, number_format="#,##0"),
    Column("py_sub", "PY subs", 8.0, number_format="#,##0"),
    Column("ac_level", "AC level", 8.0, number_format="#,##0"),
    Column("ac_from", "AC from", 8.0, number_format="#,##0"),
    Column("ac_base", "AC base", 8.0, number_format="#,##0"),
    Column("ac_add", "AC adds", 8.0, number_format="#,##0"),
    Column("ac_sub", "AC subs", 8.0, number_format="#,##0"),
    # Four series per variance tier, not two: see IMPACT_SPLIT_FORMULAS.
    Column("abs_pos_good", "Δ{ref} up ok", 7.0, number_format="+#,##0;-#,##0"),
    Column("abs_pos_bad", "Δ{ref} up bad", 7.0, number_format="+#,##0;-#,##0"),
    Column("abs_neg_good", "Δ{ref} dn ok", 7.0, number_format="+#,##0;-#,##0"),
    Column("abs_neg_bad", "Δ{ref} dn bad", 7.0, number_format="+#,##0;-#,##0"),
    Column("rel_pos_good", "Δ{ref}% up ok", 7.0, number_format=SIGNED),
    Column("rel_pos_bad", "Δ{ref}% up bad", 7.0, number_format=SIGNED),
    Column("rel_neg_good", "Δ{ref}% dn ok", 7.0, number_format=SIGNED),
    Column("rel_neg_bad", "Δ{ref}% dn bad", 7.0, number_format=SIGNED),
) + scale_columns(
    # Both waterfalls and the absolute variance are kEUR and divide by one
    # span. That is the point rather than an economy: the prior-year walk, the
    # actual walk and the variance between them are the same measure, so a
    # reader comparing a step in one to a step in the other is comparing like
    # with like only if one number scales all three.
    #
    # The walk segments print their magnitude unsigned, which is what the
    # reference prints on the deduction rows too, so their text copies carry a
    # plain format and no sign.
    (("py_base", "unit", None),
     ("py_add", "unit", "#,##0"),
     ("py_sub", "unit", "#,##0"),
     ("ac_base", "unit", None),
     ("ac_add", "unit", "#,##0"),
     ("ac_sub", "unit", "#,##0"),
     ("var_abs", "unit", "+0;-0"),
     ("abs_pos_good", "unit", None),
     ("abs_pos_bad", "unit", None),
     ("abs_neg_good", "unit", None),
     ("abs_neg_bad", "unit", None)),
    groups=("unit",)) + clip_columns(
    # The relative tier does not scale - see `clip_columns`. It clips, and the
    # text copy is what keeps the printed figure honest: the bar stops at the
    # panel edge, the label still says +983%.
    (("var_rel", None),
     ("rel_pos_good", None),
     ("rel_pos_bad", None),
     ("rel_neg_good", None),
     ("rel_neg_bad", None)),
    bounds=C12A_REL_STOP) + (
    # The relative tier's label: the measured value, and the triangles saying
    # its bar was cut. One label rather than two, because a separate marker
    # series lands at the same place as the value and the two print on top of
    # each other.
    Column("t_var_rel", "Δ{ref}% text", 8.0, number_format=None, hidden=True),
)

# The scenario-independent half of the formula set. The walk columns are added
# per row by ``waterfall_formulas`` below, because they depend on the row kind.
C12A_FORMULAS = {
    "var_abs": "={ac}{r}-{py}{r}",
    "var_rel": "=IF({py}{r}=0,NA(),({ac}{r}-{py}{r})/{py}{r}*100)",
    # What the relative tier prints: the value that was measured, and - where
    # the bar had to be cut to fit the panel - the triangles saying so. IBCS
    # varies the count with the size of the overrun rather than printing one
    # flag for everything, so a reader can tell a bar that just missed from one
    # that ran off the page. The triangles follow the number going right and
    # precede it going left, so they always point away from the axis.
    #
    # UNICHAR, not CHAR: CHAR stops at 255 and these live at U+25BA and U+25C4.
    # The whole thing is a formula, so a figure edited back inside the panel
    # loses its marker without anything being redrawn.
    "t_var_rel": (
        '=IF(ISNA({var_rel}{r}),"",'
        + f'IF({{var_rel}}{{r}}>{C12A_REL_LIMIT[1]},'
        + 'TEXT({var_rel}{r},"+0;-0")&REPT(UNICHAR(9658),'
        + f'MIN(3,MAX(1,INT({{var_rel}}{{r}}/{C12A_REL_LIMIT[1]})))),'
        + f'IF({{var_rel}}{{r}}<{C12A_REL_LIMIT[0]},'
        + f'REPT(UNICHAR(9668),MIN(3,MAX(1,INT(-{{var_rel}}{{r}}/'
        + f'{-C12A_REL_LIMIT[0]}))))&'
        + 'TEXT({var_rel}{r},"+0;-0"),'
        + 'TEXT({var_rel}{r},"+0;-0"))))'),
    **IMPACT_SPLIT_FORMULAS,
}


def waterfall_formulas(layout: "SheetLayout", template, row_index: int,
                       r: int) -> dict[str, str]:
    """The walk, for one row of one template, in both scenarios.

    Three cases, and the reason they are formulas rather than numbers is the
    whole point of the sheet: change a line and every bar below it re-floats.

        level    where the statement stands after this line
        from     where this bar starts - the previous level, zero, or the
                 level at the top of the block a group subtotal brackets
        base     the invisible segment that floats the visible one
        add/sub  the visible length, in exactly one of the two, so the fill
                 follows the sign rather than being painted on at build time
    """
    row = template.rows[row_index]
    letters = layout.letters
    out: dict[str, str] = {}
    for scen, val in (("py", "py"), ("ac", "ac")):
        level, frm = letters[f"{scen}_level"], letters[f"{scen}_from"]
        value, sign = letters[val], letters["sign"]

        if row.kind == "subtotal":
            # A subtotal states a level; it does not move it. One that also
            # advanced the total would count its own components twice.
            out[f"{scen}_level"] = (f"={level}{r - 1}" if row_index else "=0")
        else:
            prev = f"{level}{r - 1}" if row_index else "0"
            out[f"{scen}_level"] = f"={prev}+{sign}{r}*{value}{r}"

        out[f"{scen}_from"] = _c12a_start(row, r, level, row_index)
        out[f"{scen}_base"] = f"=MIN({frm}{r},{level}{r})"
        length = f"ABS({level}{r}-{frm}{r})"
        # Zero rather than NA(): these are stacked segments, and a stack needs a
        # number in every slot. NA() is right for a bar that must not exist at
        # all, which is not this case - one of the pair is always zero-length.
        out[f"{scen}_add"] = f"=IF({sign}{r}>0,{length},0)"
        out[f"{scen}_sub"] = f"=IF({sign}{r}<0,{length},0)"
    return out


C12A = SheetLayout(
    template_id="C12A",
    columns=C12A_COLUMNS,
    formulas=C12A_FORMULAS,
    n_categories=20,
    orientation="horizontal",
    chart_span=402.0,             # every panel is this tall
    # 20pt of headroom above every plot, not 8. Two things need it: the panel
    # showing the category labels would not hold the locked inset at 8 and came
    # back 0.7pt short, which verify() reported as a misaligned tier; and the
    # tier captions live in exactly this margin.
    plot_inside_lead=20.0,
    plot_inside_span=370.0,
    # Three of the four panels are kEUR, and the waterfall is what sets the
    # rate. Naming the base tier is what lets check_scale work on a template
    # with no tier called "measure".
    scale_base="wf_ac",
    same_unit_tiers=("wf_py", "wf_ac", "var_abs"),
    tiers=(
        # key,     left,  width,  bounds,        plot width, cross, plot lead
        # wf_py is wider than the others by exactly the room its category
        # labels need: it is the panel that carries the statement's line names,
        # and they are sentences rather than month abbreviations.
        TierSpec("wf_py", 0.0, 400.0, (0.0, 1120.0), 280.0, 0.0, 110.0,
                 gap_width=30, overlap=100, scale_group="unit"),
        TierSpec("wf_ac", 410.0, 300.0, (0.0, 1120.0), 280.0, 0.0, 10.0,
                 gap_width=30, overlap=100, scale_group="unit"),
        TierSpec("var_abs", 720.0, 175.0, (-65.0, 455.0), 130.0, 0.0, 20.0,
                 gap_width=30, overlap=100, caption=CAPTION_ABS,
                 scale_group="unit"),
        # No scale group. This tier clips rather than scales, for the reason
        # worked through in `clip_columns`, so its bounds stay in data units.
        TierSpec("var_rel", 905.0, 200.0, C12A_REL_LIMIT, 150.0, 0.0, 20.0,
                 gap_width=500, overlap=100, caption=CAPTION_REL),
    ),
)


# --------------------------------------------------------------------------- #
# C06F - state bars bridged to prior year by a waterfall
# --------------------------------------------------------------------------- #
#
# The first sheet whose category axis carries rows that are not categories: the
# plan, prior-year and actual totals the states sit between, and the two
# variance bars beneath them. They come from ``Template.summary_rows``, and
# ``include_summary`` is what puts them on the axis.
#
# One scale for the left panel and the bridge, and it is chosen to be exact
# rather than nearly exact: 5 kUSD per point, so the bars' 2 200 kUSD span 440pt
# and the bridge's 120pt spans exactly 600. check_scale asserts it.
#
# The bridge's axis is absolute, not a variance axis - its bars sit at levels
# between 1 600 and 2 200 - which is what lets a step be read against the bar it
# came from.

C06F_COLUMNS = (
    Column("line", "Line", 16.0, number_format=None),
    Column("kind", "Kind", 9.0, number_format=None),
    Column("scenario", "Scenario", 9.0, number_format=None),
    Column("measure", "AC", 8.0, typed=True, number_format="#,##0"),
    Column("var_rel", "Δ{ref}%", 8.0, typed=True, number_format=SIGNED),
    Column("py", "{ref}", 8.0, number_format="#,##0"),
    Column("var_abs", "Δ{ref}", 8.0, number_format="+#,##0;-#,##0"),
    # The left panel. Five series, because five different fills have to be
    # possible on one axis and Excel colours a series rather than a point.
    Column("bar_state", "AC bar", 8.0, number_format="#,##0"),
    Column("bar_short", "{ref} short", 8.0, number_format="#,##0"),
    Column("bar_pl", "PL bar", 8.0, number_format="#,##0"),
    Column("bar_py", "{ref} bar", 8.0, number_format="#,##0"),
    Column("bar_ac", "AC total", 8.0, number_format="#,##0"),
    # The bridge.
    Column("wf_level", "level", 8.0, number_format="#,##0"),
    Column("wf_from", "from", 8.0, number_format="#,##0"),
    Column("wf_base", "base", 8.0, number_format="#,##0"),
    Column("wf_up", "step up", 8.0, number_format="#,##0"),
    Column("wf_dn", "step down", 8.0, number_format="#,##0"),
    Column("rel_up", "Δ{ref}% up", 7.5, number_format=SIGNED),
    Column("rel_dn", "Δ{ref}% down", 7.5, number_format=SIGNED),
) + scale_columns(
    # `bar_short` carries no label - its number is the prior year and the
    # reference prints that only where the segment is wide enough - so it gets
    # a scaled copy and no text copy.
    #
    # This sheet is horizontal, so its relative tier is drawn by the horizontal
    # pin: an invisible `var_rel` labeller over two coloured split columns. The
    # text copy therefore belongs to `var_rel`, not to the splits.
    (("bar_state", "unit", "#,##0"),
     ("bar_short", "unit", None),
     ("bar_pl", "unit", "#,##0"),
     ("bar_py", "unit", "#,##0"),
     ("bar_ac", "unit", "#,##0"),
     ("wf_base", "unit", None),
     ("wf_up", "unit", "+#,##0"),
     ("wf_dn", "unit", '"-"#,##0'),
     ("var_rel", "rel", "+0;-0"),
     ("rel_up", "rel", None),
     ("rel_dn", "rel", None)),
    groups=("unit", "rel"))


def variance_ends(template, srow):
    """Which end of a variance bar is the reference and which is measured.

    ``span`` names two summary rows; ``reference`` names a scenario. Matching
    them by name works only while the rows happen to be called PL and PY, which
    C06F's are and C05X's ("2025 PL", "2024 AC") are not. So the reference end is
    found by asking each end what scenario it holds.
    """
    def scenario_of(label):
        row = template.summary_row(label)
        return row.stack[0][0] if row.stack else None
    ref = next((n for n in srow.span if scenario_of(n) == srow.reference),
               srow.span[0])
    measured = next(n for n in srow.span if n != ref)
    return ref, measured


def c06f_formulas(layout: "SheetLayout", template, index: int, row: int,
                  kind: str) -> dict[str, str]:
    """Every derived cell on one row of C06F, whichever kind of row it is.

    A state and a total are not the same row and cannot share one template
    string: a state's bar is its own value and its step moves the running level,
    while a total's bar is the total and its step is nothing. The variance rows
    are a third case again - they have no bar of their own and their step runs
    between two totals named by ``Summary.span``.

    Everything here references cells, so the whole sheet stays live: retype a
    state and its bar, its shortfall, its step, every level below it and the
    closing total all follow.
    """
    L = layout.letters
    out: dict[str, str] = {}

    def total_cell(label: str) -> str:
        """The measure cell of a named summary row, found by position."""
        for j, srow in enumerate(template.summary_rows):
            if srow.label == label:
                return f"${L['measure']}${layout.row_of(template, 'summary', j)}"
        raise KeyError(label)

    zeros = ("bar_state", "bar_short", "bar_pl", "bar_py", "bar_ac",
             "wf_up", "wf_dn")

    if kind == "category":
        out["py"] = (f"=IF({L['var_rel']}{row}=-100,NA(),"
                     f"{L['measure']}{row}/(1+{L['var_rel']}{row}/100))")
        out["var_abs"] = f"={L['measure']}{row}-{L['py']}{row}"
        out["bar_state"] = f"={L['measure']}{row}"
        out["bar_short"] = f"=MAX({L['py']}{row}-{L['measure']}{row},0)"
        for key in ("bar_pl", "bar_py", "bar_ac"):
            out[key] = "=0"
        # The walk. The row above is either the prior-year total, which seeds
        # the level, or another state, which has already added its own.
        out["wf_from"] = f"={L['wf_level']}{row - 1}"
        out["wf_level"] = f"={L['wf_from']}{row}+{L['var_abs']}{row}"
        out["wf_base"] = f"=MIN({L['wf_from']}{row},{L['wf_level']}{row})"
        out["wf_up"] = f"=IF({L['var_abs']}{row}>0,{L['var_abs']}{row},0)"
        out["wf_dn"] = f"=IF({L['var_abs']}{row}<0,-{L['var_abs']}{row},0)"
        out["rel_up"] = f"=IF({L['var_rel']}{row}>0,{L['var_rel']}{row},NA())"
        out["rel_dn"] = f"=IF({L['var_rel']}{row}<0,{L['var_rel']}{row},NA())"
        return out

    srow = template.summary_rows[index]
    for key in zeros:
        out[key] = "=0"
    out["py"] = "=0"
    out["rel_up"] = "=NA()"
    out["rel_dn"] = "=NA()"
    out["wf_base"] = "=0"

    if srow.is_variance_only:
        lo, hi = (total_cell(name) for name in srow.span)
        # The span says where the bar runs; the reference says which way round
        # the subtraction goes. Taking the sign from the span instead gave dPY
        # as +343 - the bar was in the right place and green, which is the kind
        # of wrong that survives a visual check.
        ref_end, measured_end = variance_ends(template, srow)
        out["var_abs"] = f"={total_cell(measured_end)}-{total_cell(ref_end)}"
        out["wf_from"] = f"={lo}"
        out["wf_level"] = f"={hi}"
        out["wf_base"] = f"=MIN({lo},{hi})"
        out["wf_up"] = f"=IF({L['var_abs']}{row}>0,{L['var_abs']}{row},0)"
        out["wf_dn"] = f"=IF({L['var_abs']}{row}<0,-{L['var_abs']}{row},0)"
        return out

    scenario = srow.stack[0][0]
    out["bar_" + ("pl" if scenario == "PL" else
                  "py" if scenario == "PY" else "ac")] = f"={L['measure']}{row}"
    # The prior-year total seeds the bridge; every other total simply carries
    # whatever level the row above reached, so the closing total shows where the
    # walk actually landed.
    out["wf_level"] = (f"={L['measure']}{row}" if scenario == "PY"
                       else ("=0" if row == layout.first_row
                             else f"={L['wf_level']}{row - 1}"))
    out["wf_from"] = f"={L['wf_level']}{row}"
    # A variance-only row already *is* its variance - it is drawn as a bar
    # between the two totals - so it does not also get a pin. Letting it have
    # one printed -17 twice, once against the actual total and once against the
    # bar that states the same thing.
    if srow.variance_rel is not None and not srow.is_variance_only:
        out["rel_up"] = f"=IF({L['var_rel']}{row}>0,{L['var_rel']}{row},NA())"
        out["rel_dn"] = f"=IF({L['var_rel']}{row}<0,{L['var_rel']}{row},NA())"
    if srow.variance_abs is not None and srow.reference:
        out["var_abs"] = (f"={L['measure']}{row}-"
                          f"{total_cell(srow.reference)}")
    return out


C06F = SheetLayout(
    template_id="C06F",
    columns=C06F_COLUMNS,
    formulas={},                  # every derived cell here is per row
    n_categories=20,              # fifteen states plus five summary rows
    orientation="horizontal",
    include_summary=True,
    chart_span=430.0,
    plot_inside_lead=20.0,
    plot_inside_span=400.0,
    scale_base="measure",
    same_unit_tiers=("measure", "wf"),
    tiers=(
        TierSpec("measure", 0.0, 560.0, (0.0, 2200.0), 440.0, 0.0, 110.0,
                 gap_width=40, overlap=100, stacked=True, scale_group="unit"),
        # An absolute axis, not a variance axis: these bars sit at levels, and
        # the window is the band the walk moves through.
        TierSpec("wf", 570.0, 140.0, (1600.0, 2200.0), 120.0, 0.0, 10.0,
                 gap_width=40, overlap=100, stacked=True, scale_group="unit"),
        TierSpec("var_rel", 720.0, 200.0, (-70.0, 40.0), 150.0, 0.0, 30.0,
                 gap_width=500, overlap=100, caption=CAPTION_REL,
                 scale_group="rel"),
    ),
)


# --------------------------------------------------------------------------- #
# C05X - monthly columns bridged to plan, with a forecast half
# --------------------------------------------------------------------------- #
#
# Three tiers down the page, seventeen categories across: the two opening
# columns, twelve months, the closing column and the two variance bars beside
# it. The measure tier and the bridge share a scale at 0.6667 kEUR per point,
# which check_scale asserts.
#
# One deliberate divergence from the reference, chosen rather than stumbled
# into. The monthly plan and actual are clustered so the plan sits behind the
# actual, and the closing column is stacked in the original - measured under
# expected. An Excel chart group is one or the other, so the workbook draws the
# closing pair side by side and the SVG keeps the stack. The alternative was to
# stack the monthly pair too, which would read as plan *plus* actual: a fidelity
# loss beats wrong notation.
#
# Both halves of the closing column, and both opening columns, are formulas over
# the months - so the sheet cannot show a total that its own rows contradict.

C05X_COLUMNS = (
    Column("period", "Period", 11.0, number_format=None),
    Column("scenario", "Scenario", 9.0, number_format=None),
    Column("kind", "Kind", 8.0, number_format=None),
    Column("measure", "AC/FC", 8.0, typed=True, number_format="#,##0"),
    Column("var_abs", "Δ{ref}", 8.0, typed=True, number_format="+#,##0;-#,##0"),
    Column("pl", "{ref}", 8.0, number_format="#,##0"),
    Column("var_rel", "Δ{ref}%", 8.5, number_format="+0.0;-0.0"),
    # The measure tier: one series per scenario, because four fills are needed
    # and Excel colours a series rather than a point.
    Column("bar_py", "PY bar", 7.5, number_format="#,##0"),
    Column("bar_pl", "{ref} bar", 7.5, number_format="#,##0"),
    Column("bar_ac", "AC bar", 7.5, number_format="#,##0"),
    Column("bar_fc", "FC bar", 7.5, number_format="#,##0"),
    # The bridge.
    Column("wf_level", "level", 7.5, number_format="#,##0"),
    Column("wf_from", "from", 7.5, number_format="#,##0"),
    Column("wf_base", "base", 7.5, number_format="#,##0"),
    # Four step series, not two: impact decides the colour and scenario decides
    # the fill, so a forecast step is hatched in its own variance colour. That
    # is four combinations, and a series can only carry one.
    Column("wf_up_ac", "up AC", 7.0, number_format="#,##0"),
    Column("wf_dn_ac", "down AC", 7.0, number_format="#,##0"),
    Column("wf_up_fc", "up FC", 7.0, number_format="#,##0"),
    Column("wf_dn_fc", "down FC", 7.0, number_format="#,##0"),
    Column("rel_up", "Δ{ref}% up", 7.5, number_format="+0.0;-0.0"),
    Column("rel_dn", "Δ{ref}% down", 7.5, number_format="+0.0;-0.0"),
    # A vertical pin's stem is a custom error bar, and an error bar takes a
    # length rather than a signed value.
    Column("rel_up_len", "up length", 7.5, number_format="0.0"),
    Column("rel_dn_len", "down length", 7.5, number_format="0.0"),
) + scale_columns(
    # The measure panel's four scenario columns and the bridge's five stacked
    # segments are all kEUR and all divide by one span - which is what keeps
    # the bridge readable against the columns beside it. `wf_base` is invisible
    # but it is a stacked segment carrying the floor, so it scales with the
    # rest or the steps land somewhere else entirely.
    #
    # The pin stems share the percentage span with the pin heads because both
    # are lengths on the same axis; scaled apart, every stem would point at
    # somewhere its head is not.
    (("bar_py", "unit", "#,##0"),
     ("bar_pl", "unit", "#,##0"),
     ("bar_ac", "unit", "#,##0"),
     ("bar_fc", "unit", "#,##0"),
     ("wf_base", "unit", None),
     ("wf_up_ac", "unit", "+#,##0"),
     ("wf_dn_ac", "unit", '"-"#,##0'),
     ("wf_up_fc", "unit", "+#,##0"),
     ("wf_dn_fc", "unit", '"-"#,##0'),
     ("rel_up", "rel", "+0.0;-0.0"),
     ("rel_dn", "rel", "+0.0;-0.0"),
     ("rel_up_len", "rel", None),
     ("rel_dn_len", "rel", None)),
    groups=("unit", "rel"))


def c05x_formulas(layout: "SheetLayout", template, index: int, row: int,
                  kind: str) -> dict[str, str]:
    """Every derived cell on one row of C05X.

    The months carry the data; the summary rows are consequences of them. Both
    opening totals and both halves of the closing column are SUMIF over the
    monthly block, so retyping a month moves the columns it belongs to - and the
    bridge, which starts on one of them and lands on another.
    """
    L = layout.letters
    out: dict[str, str] = {}
    entries = layout.sheet_entries(template)
    months = [layout.first_row + n for n, (k, _) in enumerate(entries)
              if k == "category"]
    m0, m1 = months[0], months[-1]

    def total_cell(label: str) -> str:
        for j, srow in enumerate(template.summary_rows):
            if srow.label == label:
                return f"${L['measure']}${layout.row_of(template, 'summary', j)}"
        raise KeyError(label)

    blanks = ("bar_py", "bar_pl", "bar_ac", "bar_fc")
    zeros = ("wf_up_ac", "wf_dn_ac", "wf_up_fc", "wf_dn_fc")

    if kind == "category":
        out["pl"] = f"={L['measure']}{row}-{L['var_abs']}{row}"
        out["var_rel"] = (f"=IF({L['pl']}{row}=0,NA(),"
                          f"{L['var_abs']}{row}/{L['pl']}{row}*100)")
        out["bar_py"] = "=NA()"
        out["bar_pl"] = f"={L['pl']}{row}"
        for key, scen in (("bar_ac", "AC"), ("bar_fc", "FC")):
            out[key] = (f'=IF(${L["scenario"]}{row}="{scen}",'
                        f'{L["measure"]}{row},NA())')
        out["wf_from"] = f"={L['wf_level']}{row - 1}"
        out["wf_level"] = f"={L['wf_from']}{row}+{L['var_abs']}{row}"
        out["wf_base"] = f"=MIN({L['wf_from']}{row},{L['wf_level']}{row})"
        for key, scen, test in (("wf_up_ac", "AC", ">0"), ("wf_dn_ac", "AC", "<0"),
                                ("wf_up_fc", "FC", ">0"), ("wf_dn_fc", "FC", "<0")):
            magnitude = (f"{L['var_abs']}{row}" if ">" in test
                         else f"-{L['var_abs']}{row}")
            out[key] = (f'=IF(AND(${L["scenario"]}{row}="{scen}",'
                        f'{L["var_abs"]}{row}{test}),{magnitude},0)')
        out["rel_up"] = f"=IF({L['var_rel']}{row}>0,{L['var_rel']}{row},NA())"
        out["rel_dn"] = f"=IF({L['var_rel']}{row}<0,{L['var_rel']}{row},NA())"
        out["rel_up_len"] = f"=IF({L['var_rel']}{row}>0,{L['var_rel']}{row},0)"
        out["rel_dn_len"] = f"=IF({L['var_rel']}{row}<0,-{L['var_rel']}{row},0)"
        return out

    srow = template.summary_rows[index]
    out["rel_up_len"] = "=0"
    out["rel_dn_len"] = "=0"
    for key in blanks:
        out[key] = "=NA()"
    for key in zeros:
        out[key] = "=0"
    out["pl"] = "=0"
    out["var_rel"] = "=NA()"
    out["rel_up"] = "=NA()"
    out["rel_dn"] = "=NA()"
    out["wf_base"] = "=0"
    out["wf_from"] = f"={L['wf_level']}{row}"

    if srow.is_variance_only:
        lo, hi = (total_cell(name) for name in srow.span)
        ref_end, measured_end = variance_ends(template, srow)
        out["var_abs"] = f"={total_cell(measured_end)}-{total_cell(ref_end)}"
        out["wf_from"] = f"={lo}"
        out["wf_level"] = f"={hi}"
        out["wf_base"] = f"=MIN({lo},{hi})"
        # Both closing comparisons are favourable and measured against a total,
        # so they take the solid favourable fill rather than a hatched one.
        out["wf_up_ac"] = f"=IF({L['var_abs']}{row}>0,{L['var_abs']}{row},0)"
        out["wf_dn_ac"] = f"=IF({L['var_abs']}{row}<0,-{L['var_abs']}{row},0)"
        return out

    scenario = srow.stack[0][0]
    if scenario == "PY":
        out["bar_py"] = f"={L['measure']}{row}"
        out["wf_level"] = "=0"
    elif scenario == "PL":
        out["bar_pl"] = f"={L['measure']}{row}"
        # The plan column seeds the bridge: this is where the walk starts.
        out["wf_level"] = f"={L['measure']}{row}"
    else:
        # The closing column, split by scenario over the monthly block. Side by
        # side rather than stacked - see the note at the top of this section.
        for key, scen in (("bar_ac", "AC"), ("bar_fc", "FC")):
            out[key] = (f'=SUMIF(${L["scenario"]}${m0}:${L["scenario"]}${m1},'
                        f'"{scen}",${L["measure"]}${m0}:${L["measure"]}${m1})')
        out["wf_level"] = f"={L['wf_level']}{row - 1}"
        out["var_rel"] = (f"=IF({total_cell('2025 PL')}=0,NA(),"
                          f"{L['var_abs']}{row}/{total_cell('2025 PL')}*100)")
        out["rel_up"] = f"=IF({L['var_rel']}{row}>0,{L['var_rel']}{row},NA())"
        out["rel_dn"] = f"=IF({L['var_rel']}{row}<0,{L['var_rel']}{row},NA())"
        out["rel_up_len"] = f"=IF({L['var_rel']}{row}>0,{L['var_rel']}{row},0)"
        out["rel_dn_len"] = f"=IF({L['var_rel']}{row}<0,-{L['var_rel']}{row},0)"
        out["var_abs"] = f"={L['measure']}{row}-{total_cell('2025 PL')}"
    return out


C05X = SheetLayout(
    template_id="C05X",
    columns=C05X_COLUMNS,
    formulas={},
    typed_source={"var_abs": "wf"},
    n_categories=17,
    include_summary=True,
    chart_span=620.0,
    chart_gap=18.0,
    plot_inside_lead=46.0,
    plot_inside_span=560.0,
    scale_base="measure",
    same_unit_tiers=("measure", "wf"),
    tiers=(
        TierSpec("var_rel", 10.0, 130.0, (-60.0, 80.0), 110.0, 0.0, 8.0,
                 gap_width=500, overlap=100, caption=CAPTION_REL,
                 scale_group="rel"),
        # An absolute axis: the bridge's bars sit at levels between the plan and
        # the closing column, so the window is the band the walk moves through.
        TierSpec("wf", 150.0, 110.0, (140.0, 200.0), 90.0, 0.0, 8.0,
                 gap_width=60, overlap=100, stacked=True, scale_group="unit"),
        TierSpec("measure", 270.0, 330.0, (0.0, 200.0), 300.0, 0.0, 8.0,
                 gap_width=40, overlap=76, scale_group="unit"),
    ),
)


LAYOUTS: dict[str, SheetLayout] = {"C03A": C03A, "C04A": C04A,
                                  "C05X": C05X, "C06F": C06F,
                                  "C12A": C12A}

# One registry, in the data layer.
TEMPLATES: dict[str, D.Template] = D.TEMPLATES


# --------------------------------------------------------------------------- #
# Table sheets
# --------------------------------------------------------------------------- #
#
# A table sheet drops the constraint every chart sheet is built around. There is
# no chart zone, so there is nothing to keep the data away from: the data *is*
# the presentation, and the zone rule simply does not apply.
#
# What replaces it is a different discipline, and a stricter one. On a chart
# sheet a wrong number is still a number in a cell somewhere; on a table sheet
# the cell *is* the deliverable. So every figure a reader sees is either typed
# or a formula over typed cells, every subtotal is a SUM rather than an input,
# and the threshold the red is applied at is a cell - the same cell the printed
# footnote displays. A footnote that says "-20" while the colouring uses -30 is
# the table equivalent of a total that does not follow its parts.


# --------------------------------------------------------------------------- #
# Panel geometry, shared by both renderers
# --------------------------------------------------------------------------- #
#
# Measured off the reference render in pixels, and read by the SVG renderer and
# the worksheet builder alike. It lives here rather than in either engine
# because it is the one thing they must not disagree about: the zero rule's
# position inside a panel, and the scale the panels share.
#
# The scale is per *kind*, not per panel, which is the IBCS point. Both ΔPL
# panels carry 0.4472 px per kEUR and differ in width, so a month variance of 88
# and a cumulative one of 211 are drawn to the same ruler. A panel fitted to its
# own range would make them look alike.
#
# The Excel port needs neither number in pixels - it needs the zero *fraction*
# and the value range, and both fall out of these: a panel w px wide with its
# zero z px in spans -z/scale to (w-z)/scale, whatever physical width the
# worksheet gives it.


@dataclass(frozen=True)
class PanelGeometry:
    """One drawn variance column: how wide, where its zero sits, and its scale."""

    width_px: float
    zero_px: float
    scale: float
    # Panels sharing a group share a span cell, and therefore a ruler. The two
    # dPL panels of a table are drawn at the same pixels per kEUR even though
    # one is half the width of the other - which is the whole claim of the
    # standard, and is only kept if one number drives both.
    scale_group: str | None = None

    @property
    def zero_fraction(self) -> float:
        return self.zero_px / self.width_px

    @property
    def bounds(self) -> tuple[float, float]:
        """The value-axis minimum and maximum that put the zero where it belongs."""
        return (-self.zero_px / self.scale,
                (self.width_px - self.zero_px) / self.scale)


T04A_PANELS = {
    "dpl": PanelGeometry(240, 97, 5.7143, scale_group="unit"),
    "dplp": PanelGeometry(240, 102, 3.4032, scale_group="rel"),
}

# The month block and the year-to-date block are drawn at the same scale -
# 0.4472 px per kEUR for both currency panels, 3.60 px per point for both
# percentage ones - even though the month's dPL panel is half the width. That
# is what lets a reader compare a month's miss to the year's without doing
# arithmetic, so the two share a span cell rather than each finding its own.
T02A_PANELS = {
    "dpl_november": PanelGeometry(112, 67, 0.4472, scale_group="unit"),
    "dplp_november": PanelGeometry(224, 110, 3.60, scale_group="rel"),
    "dpl_ytd_november": PanelGeometry(224, 132, 0.4472, scale_group="unit"),
    "dplp_ytd_november": PanelGeometry(224, 110, 3.60, scale_group="rel"),
}


@dataclass(frozen=True)
class TableLayout:
    """One table template's sheet.

    Row plan, top to bottom: the typed title inputs and the subject line built
    from them, a spacer, the block headers, the column captions, the data rows,
    and the threshold footnote. Every row number below is derived from that
    order rather than declared, for the same reason the chart sheets compute
    theirs - a constant that says "the data starts at row 11" stops being true
    the moment the title block gains a line.
    """

    template_id: str
    label_after_block: int = 1
    label_width: float = 17.0        # Excel character units
    value_width: float = 9.3
    row_height: float = 15.0

    # A space separates thousands, as IBCS prints them - but only above a
    # thousand. Plain "# ##0" emits the separator's space unconditionally, so
    # 59 comes out " 59" and a signed column reads "+ 59".
    measure_format: str = "[>999]# ##0{d};[<-999]-# ##0{d};0{d}"
    # A literal quoted percent sign, so the stored value is 11.8 rather than
    # 0.118. The alternative reads better as a formula but then the threshold
    # cell has to hold -0.10 while the data layer holds -10.0, and two
    # representations of one number is how a footnote starts disagreeing with
    # the colouring it describes.
    percent_format: str = '+0.0"%";-0.0"%";0.0"%"'
    # Two sections: the minus belongs inside the "<", so -20 reads "<-20" and
    # not "-<20", which is what one section produces.
    threshold_format_abs: str = '"<"0;"<-"0'
    threshold_format_rel: str = '"<"0"%";"<-"0"%"'
    # Drawn columns, keyed by tier. Empty for a table that prints every figure.
    panels: dict = field(default_factory=dict)
    # Formats for a ratio row, keyed by tier kind. Empty where a table has none.
    ratio_formats: dict = field(default_factory=dict)
    # A panel column is as many times wider than a printed one as the reference
    # draws it, so the two blocks keep the scale they share.
    panel_value_width: float = 9.3

    @property
    def title_rows(self) -> int:
        return len(TITLE_INPUTS) + 1

    def title_row(self, field: str) -> int:
        if field == "subject":
            return self.title_rows
        return TITLE_INPUTS.index(field) + 1

    # The header the reader reads, as formulas over the inputs above it. A
    # chart sheet says this in text boxes linked to the same cells; a table has
    # no shapes to link, so it says it in cells - which is simpler, and means a
    # table sheet has no way to display a title that its data does not support.
    @property
    def entity_row(self) -> int:
        return self.title_rows + 2

    @property
    def subject_row(self) -> int:
        return self.entity_row + 1

    @property
    def period_row(self) -> int:
        return self.entity_row + 2

    @property
    def block_header_row(self) -> int:
        return self.period_row + 2

    @property
    def caption_row(self) -> int:
        return self.block_header_row + 1

    @property
    def first_row(self) -> int:
        return self.caption_row + 1

    def last_row(self, template) -> int:
        return self.first_row + len(template.rows) - 1

    def footnote_row(self, template) -> int:
        return self.last_row(template) + 1

    def plan(self, template) -> list:
        return D.table_column_plan(template, self.label_after_block)

    def measure_columns(self, template) -> dict:
        """(block, scenario) -> column letter, for the variance formulas.

        A variance column has to name the two columns it is the difference of,
        and it has to find them in its own block: the cumulative dPY is the
        cumulative AC less the cumulative PY, never the month's.
        """
        out = {}
        for i, entry in enumerate(self.plan(template), start=1):
            if entry.kind == "value" and entry.tier.kind == "measure":
                out[(entry.tier.block, entry.series.scenario)] = col_letter(i)
        return out

    def columns(self, template) -> dict:
        """Excel column number for each entry of the plan, and for the labels."""
        numbers, label = {}, None
        for i, entry in enumerate(self.plan(template), start=1):
            if entry.kind == "label":
                label = i
            else:
                numbers[(entry.tier.key, entry.series_index)] = i
        return {"values": numbers, "label": label}


def subtotal_formula(template, layout: "TableLayout", index: int,
                     letter: str) -> str:
    """The SUM behind one subtotal row.

    Mirrors ``waterfall_spans``: a subtotal that spans n rows sums those n, and
    one that spans "zero" sums every element row above it. Contiguous rows of
    the same sign collapse into a single SUM range, so a country table gets
    three tidy ranges and a P&L gets the adding lines and the subtracting ones
    as separate terms - which is what makes the formula readable in the cell.
    """
    row = template.rows[index]
    skip = ("subtotal", "ratio")
    if row.spans == "zero":
        wanted = [i for i in range(index) if template.rows[i].kind not in skip]
    else:
        wanted, seen = [], 0
        for i in range(index - 1, -1, -1):
            if seen >= int(row.spans):
                break
            if template.rows[i].kind == "ratio":
                continue          # occupies a position, contributes nothing
            wanted.append(i)
            seen += 1
        wanted.reverse()

    terms, run = [], []
    def flush() -> None:
        if not run:
            return
        sign = "-" if template.rows[run[0]].sign < 0 else "+"
        first_row = layout.first_row + run[0]
        last_row = layout.first_row + run[-1]
        cells = (f"{letter}{first_row}" if first_row == last_row
                 else f"SUM({letter}{first_row}:{letter}{last_row})")
        terms.append(f"{sign}{cells}")
        run.clear()

    for i in wanted:
        if run and (template.rows[i].sign != template.rows[run[0]].sign
                    or i != run[-1] + 1):
            flush()
        run.append(i)
    flush()

    formula = "".join(terms)
    formula = formula[1:] if formula.startswith("+") else formula
    if row.spans != "zero":
        # A group subtotal states a magnitude, not a signed move: operating
        # expenses of 565.2, with the minus carried by the "-" printed against
        # the label. This is the ABS in waterfall_spans, said in a formula, so
        # the sheet and the data layer cannot disagree about the sign.
        return f"=ABS({formula})"
    return "=" + formula


T01B_TABLE = TableLayout(template_id="T01B")
T02A_TABLE = TableLayout(template_id="T02A", panels=T02A_PANELS, value_width=9.6)
T04A_TABLE = TableLayout(template_id="T04A", label_after_block=0,
                         panels=T04A_PANELS, label_width=31.0, value_width=11.5,
                         panel_value_width=11.5,
                         ratio_formats={"measure": '0.0"%"'})
T03A_TABLE = TableLayout(template_id="T03A", label_after_block=0,
                         label_width=27.0, value_width=10.0,
                         # A ratio row is denominated in something else: its
                         # figures are percentages and its absolute variance is
                         # in percentage points.
                         ratio_formats={
                             "measure": '0.0"%"',
                             "variance_abs": '+0.0"%p";-0.0"%p";0.0"%p"',
                             "variance_rel": '+0"%";-0"%";0"%"'})


# Tables are laid out by a different object, because a table sheet has no chart
# zone and none of the geometry that keeps one honest. Kept in its own registry
# so that asking for a chart layout and getting a table is impossible.
TABLE_LAYOUTS: dict[str, TableLayout] = {"T01B": T01B_TABLE,
                                         "T02A": T02A_TABLE,
                                         "T03A": T03A_TABLE,
                                         "T04A": T04A_TABLE}


def is_table(template_id: str) -> bool:
    return template_id in TABLE_LAYOUTS


def table_layout_for(template_id: str) -> TableLayout:
    try:
        return TABLE_LAYOUTS[template_id]
    except KeyError:
        raise KeyError(
            f"no table layout for {template_id!r}; have "
            f"{', '.join(sorted(TABLE_LAYOUTS))}."
        ) from None


@dataclass(frozen=True)
class StructureLayout:
    """One structure template's sheet: a data block per panel, then the charts.

    Simpler than either of the others - there are no tiers to align and no rows
    to sit a chart against - but it carries the one constraint the template
    exists to make: every panel on one scale, computed from the tallest column
    on the sheet rather than left to Excel.
    """

    template_id: str
    label_width: float = 22.0
    value_width: float = 9.0
    number_format: str = "0.0"
    max_categories: int = 5
    chart_left: float = 300.0
    chart_top: float = 10.0
    chart_height: float = 430.0
    panel_width: float = 300.0
    min_panel_width: float = 96.0
    panel_gap: float = 24.0
    gap_width: int = 44
    scale_step: float = 20.0
    chart_gap: float = 18.0
    # A band below this share of the axis gets no label. Matches the SVG's
    # pixel test, and reproduces which figures the reference leaves off.
    min_label_fraction: float = 0.03
    # Bars rather than columns. A horizontal structure chart also has to leave
    # its subtotal rows out of the chart: they are swatch-and-figure legend rows
    # and a bar chart cannot draw one.
    horizontal: bool = False

    @property
    def title_rows(self) -> int:
        return len(TITLE_INPUTS) + 1

    def title_row(self, field: str) -> int:
        if field == "subject":
            return self.title_rows
        return TITLE_INPUTS.index(field) + 1

    @property
    def first_row(self) -> int:
        return self.title_rows + 2


C01A_STRUCTURE = StructureLayout(template_id="C01A")
C02A_STRUCTURE = StructureLayout(template_id="C02A", horizontal=True,
                                 label_width=17.0, value_width=8.5,
                                 max_categories=21, number_format="# ##0",
                                 chart_height=520.0, panel_width=560.0,
                                 gap_width=26, scale_step=100.0)

STRUCTURE_LAYOUTS: dict[str, StructureLayout] = {"C01A": C01A_STRUCTURE,
                                                 "C02A": C02A_STRUCTURE}


def is_structure(template_id: str) -> bool:
    return template_id in STRUCTURE_LAYOUTS


def structure_layout_for(template_id: str) -> StructureLayout:
    try:
        return STRUCTURE_LAYOUTS[template_id]
    except KeyError:
        raise KeyError(
            f"no structure layout for {template_id!r}; have "
            f"{', '.join(sorted(STRUCTURE_LAYOUTS))}."
        ) from None



@dataclass(frozen=True)
class LineLayout:
    """One line template's sheet: series as rows, then a combo chart.

    The maximum is declared rather than fitted, for the same reason the
    structure panels share one: a line chart of cumulatives and a tier of months
    are in the same unit, and letting Excel scale to the data would put the tier
    on a different ruler from the lines it is the running sum of.
    """

    template_id: str
    label_width: float = 15.0
    value_width: float = 7.5
    number_format: str = "# ##0"
    chart_top: float = 10.0
    chart_gap: float = 18.0
    chart_width: float = 620.0
    chart_height: float = 430.0
    gap_width: int = 60
    overlap: int = 40
    maximum: float = 2200.0

    @property
    def title_rows(self) -> int:
        return len(TITLE_INPUTS) + 1

    def title_row(self, field: str) -> int:
        if field == "subject":
            return self.title_rows
        return TITLE_INPUTS.index(field) + 1

    @property
    def first_row(self) -> int:
        return self.title_rows + 2


C07C_LINE = LineLayout(template_id="C07C")
# C08H is two charts, not three tiers: the inventory level and the two movement
# series share a zero line - the increase columns rise into the stock they add
# to - so only the change tier needs an axis of its own.
C08H_LINE = LineLayout(template_id="C08H", label_width=17.0, value_width=6.2,
                       chart_width=760.0, chart_height=300.0,
                       maximum=24.0, gap_width=40, overlap=0)

LINE_LAYOUTS: dict[str, LineLayout] = {"C07C": C07C_LINE, "C08H": C08H_LINE}


def is_line(template_id: str) -> bool:
    return template_id in LINE_LAYOUTS


def line_layout_for(template_id: str) -> LineLayout:
    try:
        return LINE_LAYOUTS[template_id]
    except KeyError:
        raise KeyError(
            f"no line layout for {template_id!r}; have "
            f"{', '.join(sorted(LINE_LAYOUTS))}."
        ) from None


# --------------------------------------------------------------------------- #
# XY sheets
# --------------------------------------------------------------------------- #


@dataclass(frozen=True)
class XYLayout:
    """One XY template's sheet: a row per point, then one chart.

    The simplest data block in the project - three or four typed numbers a row,
    no formulas at all - and the most demanding chart. Everything a category
    chart gets for free has to be stated: both axis extents, both tick steps,
    and one series per scenario, because a scenario is a fill and an Excel
    series carries exactly one.

    The axis extents come from the template's own ``Axis`` objects rather than
    from anything here, which is the point: Excel left to itself picks its own
    round numbers, and a workbook whose share axis stops at 1.20 is telling a
    different story from the page that stops at 1.25.
    """

    template_id: str
    label_width: float = 12.0
    value_width: float = 10.0
    coordinate_format: str = "0.00"
    size_format: str = "0.0"
    chart_left: float = 320.0
    chart_top: float = 10.0
    chart_width: float = 660.0
    chart_height: float = 470.0
    chart_gap: float = 18.0
    # Excel scales bubbles against the largest in the chart GROUP, not the
    # largest in each series, so one group is all it takes to put every
    # scenario on one ruler. This is the percentage of the plot that largest
    # bubble fills.
    bubble_scale: int = 100
    # What makes a series. A bubble chart's points differ by scenario - prior
    # year against actual - and a scattergram's by category, the product line.
    # Both are one fill per series to Excel, and only this says which.
    series_key: str = "scenario"

    @property
    def title_rows(self) -> int:
        return len(TITLE_INPUTS) + 1

    def title_row(self, field: str) -> int:
        if field == "subject":
            return self.title_rows
        return TITLE_INPUTS.index(field) + 1

    @property
    def size_row(self) -> int:
        """Where the third measure's name and unit are typed."""
        return self.title_rows + 1

    @property
    def first_row(self) -> int:
        """The column header row; the first block of points starts below it."""
        return self.title_rows + 3

    def key_of(self, point) -> str:
        return getattr(point, self.series_key)

    def scenarios(self, template) -> list[str]:
        """The series on this chart, in the order they are painted."""
        out: list[str] = []
        for point in template.points:
            if self.key_of(point) not in out:
                out.append(self.key_of(point))
        return out

    def block(self, template, scenario: str) -> tuple[int, int]:
        """The first and last sheet row holding one scenario's points."""
        row = self.first_row + 1
        for name in self.scenarios(template):
            count = sum(1 for p in template.points if self.key_of(p) == name)
            if name == scenario:
                return row, row + count - 1
            row += count + 1          # a blank row between blocks
        raise KeyError(f"{self.template_id} has no {scenario} points")


C10D_XY = XYLayout(template_id="C10D")
# A scattergram has no size channel, so its third column is a category - the
# product line - and its series are those rather than scenarios. 149 rows in
# three blocks, so the chart wants more room than the bubble chart's twenty.
C09C_XY = XYLayout(template_id="C09C", series_key="group",
                   label_width=10.0, coordinate_format="0.00",
                   chart_left=280.0, chart_width=700.0, chart_height=560.0)

XY_LAYOUTS: dict[str, XYLayout] = {"C09C": C09C_XY, "C10D": C10D_XY}


def is_xy(template_id: str) -> bool:
    return template_id in XY_LAYOUTS


def xy_layout_for(template_id: str) -> XYLayout:
    try:
        return XY_LAYOUTS[template_id]
    except KeyError:
        raise KeyError(
            f"no XY layout for {template_id!r}; have "
            f"{', '.join(sorted(XY_LAYOUTS))}."
        ) from None



# --------------------------------------------------------------------------- #
# Tree sheets
# --------------------------------------------------------------------------- #


@dataclass(frozen=True)
class TreeLayout:
    """A driver tree's sheet: one small chart per node, plus the connectors.

    The sixth sheet family, and the one that is least like the others. A tier
    stack shares a category axis; a table shares a row grid; structure panels
    share a scale. A tree shares none of those - six charts, six zero lines,
    three different units - and what it does share is arithmetic, which is not
    a drawing property at all.

    So the geometry here is a *grid of boxes* rather than a stack of tiers, and
    the constraint the layout has to carry is the one thing IBCS requires of
    this template: boxes measuring the same unit are drawn at the same scale.
    ``px_per_unit`` therefore belongs to the scale group, not to the box, and
    ``verify_tree`` reads it back off the built charts.
    """

    template_id: str
    label_width: float = 17.0
    last_column: int = 8
    value_width: float = 8.0
    number_format: str = "0.0"
    # Excel geometry, in points. A box is a chart object; the tree columns are
    # laid left to right and the boxes stacked down each column.
    chart_left: float = 12.0
    chart_top: float = 44.0
    box_width: float = 232.0
    column_gap: float = 42.0
    row_gap: float = 9.0
    gap_width: int = 182           # narrow bars in the ratio columns
    wide_gap_width: int = 45       # the base-measure column draws them wider
    # Points per unit for each scale group. Fixed rather than fitted, because
    # a fitted axis is exactly what this template exists to argue against: let
    # Excel scale the return box to its own 5.5 maximum and it draws the same
    # height as net sales' 27.7.
    points_per_unit: dict = field(default_factory=dict)
    # Box heights, in points, keyed by node. They differ - each box is sized to
    # its own range - and that is not a contradiction of the shared scale but
    # the consequence of it.
    box_heights: dict = field(default_factory=dict)

    @property
    def title_rows(self) -> int:
        return len(TITLE_INPUTS) + 1

    def title_row(self, field_name: str) -> int:
        if field_name == "subject":
            return self.title_rows
        return TITLE_INPUTS.index(field_name) + 1

    @property
    def first_row(self) -> int:
        return self.title_rows + 2

    def scale_of(self, group: str) -> float:
        try:
            return self.points_per_unit[group]
        except KeyError:
            raise KeyError(
                f"{self.template_id} has no scale for group {group!r}; have "
                f"{', '.join(sorted(self.points_per_unit))}."
            ) from None


# Scales converted from the reference render, which is 1280px wide for a page
# drawn at 960pt: 0.75pt per px. Sharing one number per unit is the whole
# claim of the template, so they are stated once here and both engines read
# them from the data layer's measured fit.
C11A_TREE_SHEET = TreeLayout(
    template_id="C11A",
    points_per_unit={"percent": 5.3231 * 0.75,
                     "kEUR": 4.2793 * 0.75,
                     "turnover": 133.4735 * 0.75},
    box_heights={"roi": 240.0, "ros": 200.0, "turnover": 200.0,
                 "return": 96.0, "net_sales": 160.0, "capital": 144.0},
)

TREE_LAYOUTS: dict[str, TreeLayout] = {"C11A": C11A_TREE_SHEET}


def is_tree(template_id: str) -> bool:
    return template_id in TREE_LAYOUTS


def tree_layout_for(template_id: str) -> TreeLayout:
    try:
        return TREE_LAYOUTS[template_id]
    except KeyError:
        raise KeyError(
            f"no tree layout for {template_id!r}; have "
            f"{', '.join(sorted(TREE_LAYOUTS))}."
        ) from None



# --------------------------------------------------------------------------- #
# Panel sheets
# --------------------------------------------------------------------------- #


@dataclass(frozen=True)
class PanelLayout:
    """A small-multiples sheet: the grid, and the one scale every cell shares.

    The seventh sheet family, and the only one that does not draw its own
    chart. The grid is built by the `panel-charts` skill - Excel has no panel
    chart type, and the workaround of copying one chart R x C times leaves the
    user maintaining sixteen chart objects whose axes drift apart the first
    time the data changes - so this layout carries what C13 has to *tell* that
    engine, and nothing about how the engine works.
    """

    template_id: str
    grid: str = "uniform"          # which of the template's grids to build
    label_width: float = 20.0
    value_width: float = 6.0
    number_format: str = "0"
    last_column: int = 14
    chart_left: float = 12.0
    chart_top: float = 44.0
    # The band a panel's pins are drawn into, in points, and the scale that
    # fills it. Declared rather than fitted, for the reason every other sheet
    # in this workbook declares its scale: a fitted axis makes each panel its
    # own ruler, which is the one thing a small multiple may not do.
    band_points: float = 78.0
    minimum: float = -100.0
    maximum: float = 360.0

    @property
    def title_rows(self) -> int:
        return len(TITLE_INPUTS) + 1

    def title_row(self, field_name: str) -> int:
        if field_name == "subject":
            return self.title_rows
        return TITLE_INPUTS.index(field_name) + 1

    @property
    def first_row(self) -> int:
        return self.title_rows + 2


C13D_PANEL = PanelLayout(template_id="C13D")

PANEL_LAYOUTS: dict[str, PanelLayout] = {"C13D": C13D_PANEL}


def is_panel(template_id: str) -> bool:
    return template_id in PANEL_LAYOUTS


def panel_layout_for(template_id: str) -> PanelLayout:
    try:
        return PANEL_LAYOUTS[template_id]
    except KeyError:
        raise KeyError(
            f"no panel layout for {template_id!r}; have "
            f"{', '.join(sorted(PANEL_LAYOUTS))}."
        ) from None



# --------------------------------------------------------------------------- #
# Simple sheet layouts
# --------------------------------------------------------------------------- #
#
# The simple variant of a tier template is the same data drawn with fewer
# tiers, and it is expressed here rather than by cutting the Template down.
# That is deliberate and it is the cheaper half of a real trade:
#
#   * `layout.reference(t)` raises if no remaining tier names a reference
#     scenario, and every "delta-PY" column header resolves against it;
#   * `check_scale` walks `same_unit_tiers` and raises on a tier that is no
#     longer there;
#   * the tie-outs read every tier.
#
# Extra tiers on a Template are harmless - nothing draws a tier the layout does
# not list. Extra TierSpecs are fatal. So the Template keeps everything and the
# layout keeps what is drawn, which also means the simple workbook and the
# complex one are built from exactly the same numbers.
#
# The retained tier moves to offset 0 and, where the template runs
# horizontally, takes over the label margin - `add_tier` shows the category
# labels on `tiers[0]` only, so whichever tier comes first has to have room
# for them.

def _simple(base: SheetLayout, tiers: tuple[TierSpec, ...],
            same_unit: tuple[str, ...], scale_base: str) -> SheetLayout:
    """One reduced copy of a sheet layout, sharing everything not listed."""
    return replace(base, tiers=tiers, same_unit_tiers=same_unit,
                   scale_base=scale_base)


SIMPLE_LAYOUTS: dict[str, SheetLayout] = {
    # Tier 1 only: prior year behind, actual in front, forecast hatched.
    "C03A": _simple(
        C03A,
        (TierSpec("measure", 0.0, 244.0, (0.0, 230.0), 210.0, 0.0,
                  gap_width=23, overlap=76, scale_group="unit"),),
        ("measure",), "measure"),
    # Tier 1 only: nineteen states, actual against plan.
    "C04A": _simple(
        C04A,
        (TierSpec("measure", 0.0, 430.0, (0.0, 240.0), 288.0, 0.0, 110.0,
                  gap_width=20, overlap=76),),
        ("measure",), "measure"),
    # The bridge alone. C05X's subject is how the plan becomes the forecast,
    # and the waterfall is the sentence that says it - the monthly columns
    # above are the evidence, not the claim.
    "C05X": _simple(
        C05X,
        (TierSpec("wf", 0.0, 240.0, (140.0, 200.0), 200.0, 0.0, 8.0,
                  gap_width=60, overlap=100, stacked=True),),
        ("wf",), "wf"),
    # Bars and their bridge, which is the pairing the template exists to show.
    # Only the relative-variance tier goes.
    "C06F": _simple(
        C06F,
        (TierSpec("measure", 0.0, 560.0, (0.0, 2200.0), 440.0, 0.0, 110.0,
                  gap_width=40, overlap=100, stacked=True),
         TierSpec("wf", 570.0, 140.0, (1600.0, 2200.0), 120.0, 0.0, 10.0,
                  gap_width=40, overlap=100, stacked=True)),
        ("measure", "wf"), "measure"),
    # One waterfall. The actual rather than the prior year, because a statement
    # a reader is meeting for the first time should be this year's.
    "C12A": _simple(
        C12A,
        (TierSpec("wf_ac", 0.0, 400.0, (0.0, 1120.0), 280.0, 0.0, 110.0,
                  gap_width=30, overlap=100),),
        ("wf_ac",), "wf_ac"),
}

#: Which templates the simple workbook carries. The other ten have no useful
#: base-tier reduction: a scattergram and a bubble chart have no tiers at all,
#: a line chart's tiers are all measures, and a table reduced to one column
#: block is a list rather than a report.
SIMPLE_TEMPLATES = ("C01A", "C02A", "C03A", "C04A", "C05X", "C06F", "C12A")


def is_simple(template_id: str) -> bool:
    return template_id in SIMPLE_LAYOUTS


def simple_layout_for(template_id: str) -> SheetLayout:
    try:
        return SIMPLE_LAYOUTS[template_id]
    except KeyError:
        raise KeyError(
            f"no simple layout for {template_id!r}; have "
            f"{', '.join(sorted(SIMPLE_LAYOUTS))}."
        ) from None



def layout_for(template_id: str) -> SheetLayout:
    try:
        return LAYOUTS[template_id]
    except KeyError:
        raise KeyError(
            f"no Excel sheet layout for {template_id!r}; have "
            f"{', '.join(sorted(LAYOUTS))}. Add one to ibcs_layout.LAYOUTS."
        ) from None


def template_for(template_id: str) -> D.Template:
    try:
        return TEMPLATES[template_id]
    except KeyError:
        raise KeyError(
            f"unknown template {template_id!r}; have {', '.join(sorted(TEMPLATES))}"
        ) from None


if __name__ == "__main__":
    # The headers contain Δ and the Windows console defaults to cp1252, which
    # cannot encode it. Reconfiguring beats making the output ASCII.
    import sys
    sys.stdout.reconfigure(encoding="utf-8")

    for name, layout in sorted(LAYOUTS.items()):
        template = template_for(name)
        problems = layout.check_scale()
        captions = layout.check_captions(template)
        print(f"{name}: {layout.last_column} data columns "
              f"(A:{col_letter(layout.last_column)}), charts from column "
              f"{col_letter(layout.first_chart_column)} rightwards")
        print(f"  headers: {' | '.join(layout.headers(template))}")
        print(f"  rows: title 1-{layout.title_rows}, captions "
              f"{', '.join(f'{s.key}@{layout.caption_row(s)}' for s in layout.captioned)}"
              f", header {layout.header_row()}, data "
              f"{layout.first_row}-{layout.last_row()}")
        for spec in layout.captioned:
            print(f"    B{layout.caption_row(spec)}  {layout.caption_formula(spec)}"
                  f"   -> {layout.caption_text(spec, template)}")
        print(f"  category offset {layout.category_offset():.2f}pt")
        print("  scale: " + ("consistent" if not problems else "; ".join(problems)))
        print("  captions: " + ("agree with the SVG labels" if not captions
                                else "; ".join(captions)))
