"""SVG renderer for IBCS templates - the reference implementation.

Why SVG carries the notation better than a chart library: IBCS is a *geometry*
standard as much as a colour one. A relative-variance pin is a specific width
with a specific head marker; tiers share one category axis exactly; the variance
axis is a double rule of a specific gap. All of that is arithmetic here, and a
fight with defaults anywhere else.

This module is also where each template's geometry gets solved first. The Excel
builder reproduces what this produces, rather than re-deriving it, which is what
keeps the two renderers agreeing.

Layout numbers are measured off the IBCS reference renders in template-refs/,
not invented, so a diff against the original is meaningful. Measuring is done
ad hoc per template - scan the PNG for the bar bands, the axis rules and the
group rules, then solve the scale from two known values at opposite ends of the
range - and the numbers land in that template's ``*_LAYOUT`` dict with a comment
saying what they were taken from.
"""

from __future__ import annotations

from dataclasses import dataclass
from math import sqrt
from typing import Sequence
from xml.sax.saxutils import escape

import ibcs_data as D
import ibcs_layout as L
import ibcs_style as S
import ibcs_paths as P
# Panel geometry is measured once and read by both renderers; see
# ibcs_layout.PanelGeometry for why it lives there rather than here.
from ibcs_layout import T02A_PANELS, T04A_PANELS

FONT = "Arial, Helvetica, sans-serif"


# --------------------------------------------------------------------------- #
# Geometry
# --------------------------------------------------------------------------- #


@dataclass(frozen=True)
class CategoryAxis:
    """Horizontal band positions shared by every tier in a template.

    This single object is the whole answer to the multi-tier alignment problem.
    Every tier asks it for the same x for a given category, so tiers cannot
    drift; there is no per-chart plot area to keep in sync.
    """

    x0: float
    pitch: float
    count: int

    @property
    def x1(self) -> float:
        return self.x0 + self.pitch * self.count

    def centre(self, index: int) -> float:
        return self.x0 + self.pitch * (index + 0.5)

    def span(self, index: int) -> tuple[float, float]:
        left = self.x0 + self.pitch * index
        return left, left + self.pitch


@dataclass(frozen=True)
class Scale:
    """Maps a data value to a y coordinate within one tier.

    ``units_per_px`` is deliberately explicit rather than derived from the data
    range: IBCS requires tiers sharing a unit to share a scale, and that only
    happens if the caller can pass the same number to both.
    """

    zero_y: float
    px_per_unit: float

    def y(self, value: float) -> float:
        return self.zero_y - value * self.px_per_unit

    def height(self, value: float) -> float:
        return abs(value) * self.px_per_unit


# --------------------------------------------------------------------------- #
# Canvas
# --------------------------------------------------------------------------- #


class Canvas:
    """Accumulates SVG elements and the pattern definitions they reference.

    ``prefix`` namespaces every generated id. This is not tidiness: SVG ids are
    global to the document, so inlining several charts into one page - a gallery
    grid, say - makes identically-named patterns collide, and the later chart
    silently repaints the earlier one's forecast bars. The failure looks like a
    rendering quirk rather than an id clash, so it is worth preventing outright.
    """

    def __init__(self, width: float, height: float, prefix: str = "ibcs") -> None:
        self.width = width
        self.height = height
        self.prefix = prefix
        self._defs: dict[str, str] = {}
        self._body: list[str] = []

    # -- primitives ------------------------------------------------------- #

    def add(self, markup: str) -> None:
        self._body.append(markup)

    def rect(self, x, y, w, h, fill="none", stroke=None, stroke_width=1, **kw) -> None:
        if w <= 0 or h <= 0:
            return
        attrs = f'x="{x:.2f}" y="{y:.2f}" width="{w:.2f}" height="{h:.2f}" fill="{fill}"'
        if stroke:
            attrs += f' stroke="{stroke}" stroke-width="{stroke_width}"'
        attrs += self._extra(kw)
        self.add(f"<rect {attrs}/>")

    def line(self, x1, y1, x2, y2, stroke="#000000", stroke_width=1, **kw) -> None:
        self.add(
            f'<line x1="{x1:.2f}" y1="{y1:.2f}" x2="{x2:.2f}" y2="{y2:.2f}" '
            f'stroke="{stroke}" stroke-width="{stroke_width}"{self._extra(kw)}/>'
        )

    def ellipse(self, cx, cy, rx, ry, stroke, stroke_width=1.5, fill="none", **kw) -> None:
        self.add(
            f'<ellipse cx="{cx:.2f}" cy="{cy:.2f}" rx="{rx:.2f}" ry="{ry:.2f}" '
            f'fill="{fill}" stroke="{stroke}" stroke-width="{stroke_width}"{self._extra(kw)}/>'
        )

    def text(self, x, y, content, size=11, fill="#000000", anchor="start",
             weight="normal", style="normal", **kw) -> None:
        self.add(
            f'<text x="{x:.2f}" y="{y:.2f}" font-family="{FONT}" font-size="{size}" '
            f'fill="{fill}" text-anchor="{anchor}" font-weight="{weight}" '
            f'font-style="{style}"{self._extra(kw)}>{escape(str(content))}</text>'
        )

    def rich_text(self, x, y, runs: Sequence[tuple[str, str]], size=11,
                  fill="#000000") -> None:
        """One line of text mixing weights, e.g. a bold comment lead then body.

        Emitted as tspans inside a single text element so the browser does the
        kerning; positioning each run by estimated width would drift.
        """
        spans = "".join(
            f'<tspan font-weight="{weight}">{escape(content)}</tspan>'
            for content, weight in runs
        )
        self.add(
            f'<text x="{x:.2f}" y="{y:.2f}" font-family="{FONT}" font-size="{size}" '
            f'fill="{fill}" xml:space="preserve">{spans}</text>'
        )

    @staticmethod
    def _extra(kw: dict) -> str:
        return "".join(f' {k.replace("_", "-")}="{v}"' for k, v in kw.items())

    # -- patterns --------------------------------------------------------- #

    def hatch(self, hatch: S.Hatch, background: str) -> str:
        """Register (once) the FC diagonal hatch and return its paint reference.

        The palette records the stripe geometry as measured along a horizontal
        scanline of the reference render. A 45-degree stripe crossed horizontally
        appears sqrt(2) wider than it truly is, so both pitch and stroke are
        divided by sqrt(2) to get the perpendicular values SVG wants.
        """
        pid = (f"{self.prefix}-hatch-{hatch.colour.lstrip('#')}"
               f"-{background.lstrip('#')}")
        if pid not in self._defs:
            pitch = hatch.pitch_px / sqrt(2)
            stroke = hatch.stroke_px / sqrt(2)
            rotate = 45 if hatch.direction == "ascending" else -45
            self._defs[pid] = (
                f'<pattern id="{pid}" width="{pitch:.3f}" height="{pitch:.3f}" '
                f'patternUnits="userSpaceOnUse" patternTransform="rotate({rotate})">'
                f'<rect width="{pitch:.3f}" height="{pitch:.3f}" fill="{background}"/>'
                f'<rect width="{stroke:.3f}" height="{pitch:.3f}" fill="{hatch.colour}"/>'
                f"</pattern>"
            )
        return f"url(#{pid})"

    # -- output ----------------------------------------------------------- #

    def to_svg(self) -> str:
        """Emit the document.

        width/height and viewBox are both present on purpose: the explicit size
        gives the file an intrinsic dimension for a bare <img>, while the viewBox
        lets it scale to whatever cell a grid layout puts it in. Dropping either
        breaks one of the two cases.
        """
        defs = "".join(self._defs.values())
        return (
            f'<svg xmlns="http://www.w3.org/2000/svg" width="{self.width:.0f}" '
            f'height="{self.height:.0f}" viewBox="0 0 {self.width:.0f} {self.height:.0f}" '
            f'preserveAspectRatio="xMidYMid meet" role="img">'
            f"<defs>{defs}</defs>"
            f'<rect width="{self.width:.0f}" height="{self.height:.0f}" '
            f'fill="{S.PAGE["background"]}"/>'
            + "".join(self._body)
            + "</svg>"
        )


# --------------------------------------------------------------------------- #
# Notation primitives
# --------------------------------------------------------------------------- #


def paint_for(canvas: Canvas, scenario: str) -> tuple[str, str | None]:
    """Return (fill_paint, outline) for a scenario, registering a pattern if needed."""
    spec = S.scenario_fill(scenario)
    if spec.hatch:
        return canvas.hatch(spec.hatch, spec.fill), spec.outline
    return spec.fill, spec.outline


def scenario_bar(canvas: Canvas, scenario: str, cx: float, width: float,
                 y_value: float, y_zero: float) -> None:
    """A column drawn in one scenario's notation."""
    fill, outline = paint_for(canvas, scenario)
    top, bottom = sorted((y_value, y_zero))
    canvas.rect(cx - width / 2, top, width, bottom - top,
                fill=fill, stroke=outline, stroke_width=1 if outline else 0)


def variance_pin(canvas: Canvas, scenario: str, cx: float, stem_width: float,
                 y_value: float, y_zero: float, colour: str,
                 head: float | None = None,
                 head_height: float | None = None) -> None:
    """A relative-variance pin: a thin stem in the variance colour, with a head
    marker at the value end carrying the *minuend's* scenario fill.

    That head marker is the part people drop, and it is load-bearing: it is what
    tells the reader whether a relative variance was measured or forecast.

    The head is emitted *before* the stem, so the stem reads its full length
    across it. Drawn the other way the head covers the last few pixels of the
    stem and the pin stops short of the value it marks - which is what the
    reference shows it should not do. Same rectangles either way; only the
    order changes, so no geometry moves.
    """
    # Square by default, from the palette's ratio. C13 passes both dimensions
    # because its heads are measurably wider than tall - 8 by 7 against a 3px
    # stem - which a single ratio cannot express.
    head = stem_width * S.PIN_HEAD_RATIO if head is None else head
    tall = head if head_height is None else head_height
    fill, outline = paint_for(canvas, scenario)
    canvas.rect(cx - head / 2, y_value - tall / 2, head, tall,
                fill=fill, stroke=outline or fill, stroke_width=0.8)

    top, bottom = sorted((y_value, y_zero))
    canvas.rect(cx - stem_width / 2, top, stem_width, bottom - top, fill=colour)


def reference_axis(canvas: Canvas, scenario: str, x0: float, x1: float, y: float) -> None:
    """Draw a variance tier's axis in the notation of its reference scenario.

    Solid light for a PY comparison; two thin parallel rules for PL or BU, which
    is the outlined 'fictitious' fill seen edge-on. This is why a correct
    variance tier needs no legend.
    """
    spec = S.reference_axis(scenario)
    if spec["style"] == "solid":
        canvas.rect(x0, y - spec["weight_px"] / 2, x1 - x0, spec["weight_px"],
                    fill=spec["colour"])
    else:
        gap = spec["gap_px"] / 2 + spec["weight_px"] / 2
        for offset in (-gap, gap):
            canvas.line(x0, y + offset, x1, y + offset,
                        stroke=spec["colour"], stroke_width=spec["weight_px"])


# --- horizontal counterparts ---------------------------------------------- #
#
# Columns run up from a category axis along the bottom; bars run right from one
# down the side. IBCS treats that as a semantic choice - time goes across,
# structure goes down - so a template declares its orientation and the renderer
# needs both code paths. Each function below is the transpose of the one above
# it: the value axis becomes x, the category axis becomes y.


def scenario_bar_h(canvas: Canvas, scenario: str, cy: float, height: float,
                   x_value: float, x_zero: float) -> None:
    """A bar drawn in one scenario's notation."""
    fill, outline = paint_for(canvas, scenario)
    left, right = sorted((x_value, x_zero))
    canvas.rect(left, cy - height / 2, right - left, height,
                fill=fill, stroke=outline, stroke_width=1 if outline else 0)


def variance_pin_h(canvas: Canvas, scenario: str, cy: float, stem_height: float,
                   x_value: float, x_zero: float, colour: str) -> None:
    """A relative-variance pin lying on its side.

    The head still carries the minuend's scenario fill - that is what says
    whether the variance was measured or forecast - and still sits at the value
    end, not the axis end.
    """
    head = stem_height * S.PIN_HEAD_RATIO
    fill, outline = paint_for(canvas, scenario)
    canvas.rect(x_value - head / 2, cy - head / 2, head, head,
                fill=fill, stroke=outline or fill, stroke_width=0.8)

    left, right = sorted((x_value, x_zero))
    canvas.rect(left, cy - stem_height / 2, right - left, stem_height, fill=colour)


# --------------------------------------------------------------------------- #
# Lines
# --------------------------------------------------------------------------- #
#
# IBCS insists a line chart shows its markers, and the reason is worth keeping:
# the line between two points is a connector, not data. Business measures are
# not continuous - there is no value at half past March - so a line drawn
# without markers invites a reader to take one.
#
# The marker is also where the scenario lives. A line has no fill to carry it,
# so each point takes the marker of its own scenario: solid for an actual,
# hollow for a plan, hatched for a forecast. That is what lets one line change
# from measured to expected part way along, which is exactly what C07 does in
# September.


def scenario_line(canvas: Canvas, points: Sequence[tuple[float, float]],
                  colour: str, width: float = 1.0) -> None:
    """The connector between markers. Thin, and drawn under them."""
    if len(points) < 2:
        return
    d = " ".join(("M" if i == 0 else "L") + f" {x:.2f} {y:.2f}"
                 for i, (x, y) in enumerate(points))
    canvas.add(f'<path d="{d}" fill="none" stroke="{colour}" '
               f'stroke-width="{width:.2f}"/>')


def scenario_marker(canvas: Canvas, scenario: str, cx: float, cy: float,
                    size: float) -> None:
    """A square marker in its scenario's notation.

    Same three statements a bar makes with its fill, at the size of a marker -
    which is the whole point: a reader who has learned the fills reads a line
    chart without being taught anything new.
    """
    fill, outline = paint_for(canvas, scenario)
    canvas.rect(cx - size / 2, cy - size / 2, size, size, fill=fill,
                stroke=outline or S.TEXT["primary"], stroke_width=0.9)

def reference_axis_v(canvas: Canvas, scenario: str, y0: float, y1: float,
                     x: float) -> None:
    """A vertical variance axis in the notation of its reference scenario.

    The double rule is the whole point of this function. PL and BU are
    *fictitious* scenarios, drawn outlined rather than filled, and an outlined
    bar seen edge-on is two parallel lines - so the axis of a dPL tier states
    what the tier is measured against without a legend. Drawing it as a single
    line, which is all a chart library will give you, throws that away.
    """
    spec = S.reference_axis(scenario)
    if spec["style"] == "solid":
        canvas.rect(x - spec["weight_px"] / 2, y0, spec["weight_px"], y1 - y0,
                    fill=spec["colour"])
    else:
        gap = spec["gap_px"] / 2 + spec["weight_px"] / 2
        for offset in (-gap, gap):
            canvas.line(x + offset, y0, x + offset, y1,
                        stroke=spec["colour"], stroke_width=spec["weight_px"])


def highlight_oval(canvas: Canvas, cx: float, cy: float, rx: float, ry: float,
                   ref: int | None = None) -> None:
    """Blue ellipse round a value, optionally with its comment reference circle."""
    blue = S.ANNOTATION["highlight"]
    canvas.ellipse(cx, cy, rx, ry, stroke=blue, stroke_width=1.4)
    if ref is not None:
        rcx = cx + rx + 11
        canvas.ellipse(rcx, cy, 8, 8, stroke=blue, stroke_width=1.2)
        canvas.text(rcx, cy + 3.5, ref, size=10, fill=blue, anchor="middle")


# --------------------------------------------------------------------------- #
# C03A layout - measured from template-refs/C03_03A.png
# --------------------------------------------------------------------------- #

C03A_LAYOUT = dict(
    page=(1280, 720),
    margin_left=20,
    margin_right=20,
    cat_x0=64,
    cat_pitch=56,
    bar_width=38,
    ac_offset=9,            # AC sits right of the band centre; PY sits on it
    pin_stem=5,
    title_rule_y=80,
    tier_rel_zero=160,
    tier_rel_scale=2.59,    # px per percentage point
    tier_abs_zero=304,
    measure_zero=639,
    unit_scale=1.320,       # px per kEUR - shared by the measure and dPY tiers
    split_after=9,          # vertical rule between the last AC and first FC period
    split_top=95,
    split_bottom=678,
    summary_cx=827,         # AC bar centre of the detached year block
    summary_py_dx=-22,
    summary_width=44,
    summary_scale=0.1110,   # the year block is compressed to fit; see note below
    index_value=100,        # the scale-break reference band, in measure units
    index_strip=7,
    comments_x=988,
    comments_y=(490, 580),
    # Type sizes measured off the reference render. IBCS keeps the value labels
    # and the category labels at the same size as the body text: the numbers are
    # the content, not a caption on it.
    font=dict(entity=13, measure=13.5, period=13, message=13.5, badge=34,
              value=12.5, category=13, tier=12.5, comment=12.5, footer=9),
)


def render_c03a(template: D.Template = D.C03A, layout: dict | None = None) -> str:
    """Render C03A - Furniture Inc. contribution, three tiers plus a year block."""
    L = {**C03A_LAYOUT, **(layout or {})}
    width, height = L["page"]
    c = Canvas(width, height, prefix=f"{template.id}{template.variant}".lower())

    cats = CategoryAxis(L["cat_x0"], L["cat_pitch"], len(template.categories))
    rel = Scale(L["tier_rel_zero"], L["tier_rel_scale"])
    absv = Scale(L["tier_abs_zero"], L["unit_scale"])
    meas = Scale(L["measure_zero"], L["unit_scale"])

    scen = template.category_scenarios
    bar_w = L["bar_width"]
    ac_dx = L["ac_offset"]
    F = L["font"]

    _title_block(c, template, L)

    # --- tier 1: relative variance (pins) -------------------------------- #
    tier = template.tier("var_rel")
    reference_axis(c, tier.reference, cats.x0, cats.x1, rel.zero_y)
    for i, sc in enumerate(scen):
        value = _value(tier, sc, i)
        if value is None:
            continue
        cx = cats.centre(i) + ac_dx
        colour = S.variance_colour(value, tier.higher_is_better)
        variance_pin(c, sc, cx, L["pin_stem"], rel.y(value), rel.zero_y, colour)
        _stacked_label(c, cx, rel.y(value), value, tier.label_for(i, value),
                       head=L["pin_stem"] * S.PIN_HEAD_RATIO, size=F["value"])
    c.text(cats.x0 - 6, rel.zero_y + 4, tier.label, size=F["tier"], anchor="end")

    # --- tier 2: absolute variance (columns) ----------------------------- #
    tier = template.tier("var_abs")
    reference_axis(c, tier.reference, cats.x0, cats.x1, absv.zero_y)
    for i, sc in enumerate(scen):
        value = _value(tier, sc, i)
        if value is None:
            continue
        cx = cats.centre(i) + ac_dx
        colour = S.variance_colour(value, tier.higher_is_better)
        _variance_column(c, sc, cx, bar_w, absv.y(value), absv.zero_y, colour)
        _stacked_label(c, cx, absv.y(value), value, tier.label_for(i, value),
                       size=F["value"])
    c.text(cats.x0 - 6, absv.zero_y + 4, tier.label, size=F["tier"], anchor="end")

    # --- tier 3: the measure --------------------------------------------- #
    tier = template.tier("measure")
    py = tier.series_for("PY")
    _index_band(c, cats, meas, L)
    for i in range(len(scen)):
        if py.values[i] is not None:
            scenario_bar(c, "PY", cats.centre(i), bar_w, meas.y(py.values[i]), meas.zero_y)
    for i, sc in enumerate(scen):
        value = _value(tier, sc, i)
        if value is None:
            continue
        cx = cats.centre(i) + ac_dx
        scenario_bar(c, sc, cx, bar_w, meas.y(value), meas.zero_y)
        c.text(cx, meas.y(value) - 6, _tier_number(tier, i, value),
               size=F["value"], anchor="middle")

    c.line(cats.x0, meas.zero_y, cats.x1, meas.zero_y, stroke="#000000", stroke_width=1.6)

    # Scenario labels sit beside the first bars, which is IBCS's integrated
    # legend: name the series where it is, not in a box somewhere else. AC is
    # levelled with January's value label so the two read as one line.
    # From the template's own first actual, not from C03A's. Levelling the
    # caption with a value transcribed from the reference put "AC" halfway up
    # somebody else's chart the moment the numbers changed.
    _ac_series = tier.series_for("AC")
    _first_ac = next((v for v in (_ac_series.values if _ac_series
                                  else ()) if v is not None), 0.0)
    c.text(cats.x0 - 6, meas.y(_first_ac) - 6, "AC", size=F["value"], anchor="end")
    # A layout key rather than a constant. 76pt above the zero is where the
    # prior-year bars sit on IBCS's numbers; on a chart whose bars are a
    # tenth the height it lands on top of the AC caption.
    c.text(cats.x0 - 6, meas.zero_y - L.get("py_label_lift", 76), "PY",
           size=F["value"], anchor="end")

    _category_labels(c, cats, template, L, meas.zero_y)
    _period_split(c, cats, L)
    _summary_block(c, template, L, meas.zero_y,
                   cats.x1 + L["index_strip"],
                   meas.y(L["index_value"]) if L.get("index_value") is not None
                   else meas.zero_y)
    _annotations(c, template, cats, L, rel, absv, meas)
    _comments(c, template, L)
    _footer(c, L)

    return c.to_svg()


# --------------------------------------------------------------------------- #
# C03A parts
# --------------------------------------------------------------------------- #


def _value(tier: D.Tier, scenario: str, index: int) -> float | None:
    series = tier.series_for(scenario)
    return None if series is None else series.values[index]


def _variance_column(canvas: Canvas, scenario: str, cx: float, width: float,
                     y_value: float, y_zero: float, colour: str) -> None:
    """Absolute variances are full-width columns in the variance colour.

    A forecast variance keeps the hatch, so the reader can see that the variance
    itself is expected rather than measured - drawn as hatch in the variance
    colour, not in the scenario grey.
    """
    top, bottom = sorted((y_value, y_zero))
    if scenario == "FC":
        hatch = S.scenario_fill("FC").hatch
        paint = canvas.hatch(S.Hatch(colour, hatch.angle_deg, hatch.direction,
                                     hatch.pitch_px, hatch.stroke_px), "#FFFFFF")
        canvas.rect(cx - width / 2, top, width, bottom - top,
                    fill=paint, stroke=colour, stroke_width=1)
    else:
        canvas.rect(cx - width / 2, top, width, bottom - top, fill=colour)


def _stacked_label(canvas: Canvas, cx: float, y_value: float, value: float,
                   text: str, head: float = 0.0, size: float = 12.5) -> None:
    """Value labels sit outside the element, in the direction it points."""
    offset = head / 2 + 5
    y = y_value - offset - 2 if value >= 0 else y_value + offset + size * 0.8
    canvas.text(cx, y, text, size=size, anchor="middle")


def _index_band(canvas: Canvas, cats: CategoryAxis, scale: Scale, L: dict) -> None:
    """The scale-break reference: a 0-to-N band drawn only in the gaps between bars.

    This is how C03A stays honest about the year block having a different scale.
    The same 100 kEUR is drawn at both scales, so a reader can see the
    compression rather than having to infer it from the axis - which is the only
    thing that makes a broken scale acceptable at all.
    """
    # A chart with no detached summary block has no broken scale, and the band
    # exists only to make a break honest. Drawn anyway it is decoration, and on
    # a three-category chart it is decoration in the wrong places.
    if L.get("index_value") is None:
        return
    pale = S.ANNOTATION["highlight_pale"]
    top = scale.y(L["index_value"])
    height = scale.zero_y - top
    for i in range(cats.count + 1):
        x = cats.x0 + cats.pitch * i + 1
        canvas.rect(x, top, L["index_strip"], height, fill=pale)
    size = L["font"]["category"]
    canvas.text(cats.x0 - 6, top + 5, f"{L['index_value']:.0f}", size=size, anchor="end")
    canvas.text(cats.x1 + 16, top + 5, f"{L['index_value']:.0f}", size=size)


def _title_block(canvas: Canvas, t: D.Template, L: dict) -> None:
    x, F = L["margin_left"], L["font"]
    canvas.text(x, L.get("entity_baseline", 22), t.title.entity,
                size=F["entity"])
    # Measure name bold, unit in secondary grey: the subject of the chart and
    # what it is counted in are different kinds of fact. Emitted as one text
    # element so the space between them survives regardless of width estimates.
    # The suffix, where a template has one, states what the ordering means -
    # C04A is sorted by dPL and that ordering *is* its message, so leaving it
    # unstated would hide the reason the rows are where they are.
    # A chart with no single unit prints none. C10 plots three measures in three
    # different units, so the subject line names the subject and each unit is
    # stated where it is used - two axis titles and a size legend.
    suffix = (f" in {t.title.unit}" if t.title.unit else "")
    runs = list(t.title.subject) or [(t.title.measure, "bold"), (suffix, "normal")]
    runs = runs[:-1] + [(runs[-1][0] + L.get("subject_suffix", ""), runs[-1][1])]
    canvas.rich_text(x, L.get("subject_baseline", 41), runs, size=F["measure"])
    canvas.text(x, L.get("period_baseline", 59), t.title.period, size=F["period"])

    # The message is the actual title - a sentence stating the point, not a label.
    width = L.get("message_width", 545)
    for n, line in enumerate(_wrap_runs([(t.title.message, "normal")], width,
                                        F["message"])):
        canvas.rich_text(L.get("message_x", 400),
                         L.get("message_top", 22)
                         + n * L.get("message_leading", 17),
                         line, size=F["message"])

    # The template number, which only means something on a recreation of that
    # template. A chart of somebody else's data passes badge="" and gets none -
    # printing "12A" over a company's own P&L claims it is IBCS's example.
    badge = L.get("badge")
    if badge is None:
        stem = t.id[1:] if t.id.startswith("C") else t.id
        badge = f"{stem}{t.variant}"
    if badge:
        canvas.text(L.get("badge_right", L["page"][0] - L["margin_right"]),
                    L.get("badge_baseline", 48), badge,
                    size=F["badge"], weight="bold", fill=S.TEXT["badge"],
                    anchor="end")
    canvas.line(x, L["title_rule_y"], L["page"][0] - L["margin_right"],
                L["title_rule_y"], stroke=S.PAGE["rule"], stroke_width=1)


def _category_labels(canvas: Canvas, cats: CategoryAxis, t: D.Template,
                     L: dict, baseline: float) -> None:
    size = L["font"]["category"]
    for i, name in enumerate(t.categories):
        canvas.text(cats.centre(i) + L["ac_offset"], baseline + 19, name,
                    size=size, anchor="middle")
    # The scenario of a period is stated once, under the first period it applies
    # to - the same economy as the integrated legend.
    canvas.text(cats.centre(0) + L["ac_offset"], baseline + 34, "AC",
                size=size, anchor="middle")
    after = L.get("split_after")
    if after is not None and after < len(t.categories):
        canvas.text(cats.centre(after) + L["ac_offset"], baseline + 34, "FC",
                    size=size, anchor="middle")


def _period_split(canvas: Canvas, cats: CategoryAxis, L: dict) -> None:
    """The rule between measured and expected periods.

    IBCS wants the boundary between what happened and what is expected to be
    unmissable; a change of fill alone is not enough at a glance.
    """
    after = L.get("split_after")
    # Nothing to separate when every period is measured. A rule drawn past the
    # last category is a rule about nothing, and the "FC" caption under it
    # names a scenario the chart does not carry.
    if after is None or after >= cats.count:
        return
    x = cats.x0 + cats.pitch * after + 0.5
    canvas.line(x, L["split_top"], x, L["split_bottom"], stroke="#000000", stroke_width=1.2)


def _summary_block(canvas: Canvas, t: D.Template, L: dict, baseline: float,
                   month_index_x: float, month_index_y: float) -> None:
    """The detached full-year block.

    It carries its own, compressed scale because an annual total cannot share a
    monthly scale and stay on the page. IBCS allows this only when the break is
    made visible - hence the gap, and the 100-index markers on both scales.
    """
    s = t.summary
    if s is None:
        return
    cx, w, F = L["summary_cx"], L["summary_width"], L["font"]
    scale = Scale(baseline, L["summary_scale"])
    left, right = cx - w / 2, cx + w / 2

    # The same 0-to-100 band as the monthly plot, at this block's own scale, with
    # a connector back to the monthly one. Seeing the same 100 kEUR at two
    # heights is what makes a broken scale honest rather than misleading.
    pale = S.ANNOTATION["highlight_pale"]
    idx_top = scale.y(L["index_value"])
    strip_x = left + L["summary_py_dx"] - 12
    for x in (strip_x, right + 2):
        canvas.rect(x, idx_top, L["index_strip"] - 2, baseline - idx_top, fill=pale)
    canvas.text(right + 14, idx_top + 5, f"{L['index_value']:.0f}", size=F["category"])
    canvas.line(month_index_x, month_index_y, strip_x, idx_top,
                stroke=S.ANNOTATION["highlight"], stroke_width=1)

    py_year = sum(v for _, v in s.stack) / (1 + s.variance_rel / 100)
    scenario_bar(canvas, "PY", cx + L["summary_py_dx"], w, scale.y(py_year), baseline)

    cursor = 0.0
    for scenario, value in s.stack:
        scenario_bar(canvas, scenario, cx, w, scale.y(cursor + value), scale.y(cursor))
        mid = (scale.y(cursor) + scale.y(cursor + value)) / 2
        canvas.text(cx, mid + 4, _thousands(value), size=F["value"], anchor="middle",
                    fill="#FFFFFF" if scenario == "AC" else "#000000")
        canvas.text(right + 5, mid + 4, scenario, size=F["value"])
        cursor += value

    total = sum(v for _, v in s.stack)
    canvas.text(cx, scale.y(total) - 7, _thousands(total), size=F["value"], anchor="middle")
    canvas.text(cx, baseline + 19, s.label, size=F["category"], anchor="middle")

    # The year block's variance tiers sit on the same tier baselines as the
    # months, so the reader's eye does not have to re-find the zero line - but
    # each gets its own short axis segment, because the block is detached.
    axis_x0, axis_x1 = left + L["summary_py_dx"] - 4, right + 4
    for zero, value, stem in (
        (L["tier_abs_zero"], s.variance_abs, None),
        (L["tier_rel_zero"], s.variance_rel, L["pin_stem"]),
    ):
        reference_axis(canvas, s.reference, axis_x0, axis_x1, zero)
        sc = Scale(zero, L["unit_scale"] if stem is None else L["tier_rel_scale"])
        colour = S.variance_colour(value, True)
        fmt = "{:+,.0f}" if stem is None else "{:+.1f}"
        if stem is None:
            _variance_column(canvas, "FC", cx, w, sc.y(value), zero, colour)
        else:
            variance_pin(canvas, "FC", cx, stem, sc.y(value), zero, colour)
        _stacked_label(canvas, cx, sc.y(value), value, fmt.format(value),
                       head=0 if stem is None else stem * S.PIN_HEAD_RATIO,
                       size=F["value"])


def _thousands(value: float) -> str:
    """IBCS separates thousands with a thin space, not a comma."""
    return f"{value:,.0f}".replace(",", "\u2009")



def _tier_number(tier: D.Tier, index: int, value: float) -> str:
    """A tier's own number format, with IBCS's thousands separator.

    `_thousands` fixes the decimals at none, which is right for a
    template reporting kEUR in the hundreds and wrong for one
    reporting cash in the tens - 4.295 came out as "4". The tier
    already declares how many decimals its numbers carry, so ask it,
    and then apply the separator IBCS wants.
    """
    return tier.label_for(index, value).replace(",", " ")


def _annotations(canvas: Canvas, t: D.Template, cats: CategoryAxis, L: dict,
                 rel: Scale, absv: Scale, meas: Scale) -> None:
    for a in t.annotations:
        tier_key, index = a.target
        if tier_key == "var_abs" and index == -1:
            highlight_oval(canvas, L["summary_cx"] - 8,
                           absv.y(t.summary.variance_abs) - 14, 22, 11, a.ref)
        elif tier_key == "measure":
            # Derived, for the same reason: an oval is placed on the value the
            # message names, and that value belongs to this template. By the
            # period's *own* scenario, not by AC - C03A's oval sits on a
            # December forecast, where the actual series holds nothing.
            measure = t.tier("measure")
            value = _value(measure, t.category_scenarios[index], index)
            if value is None:
                continue
            cx = cats.centre(index) + L["ac_offset"]
            highlight_oval(canvas, cx, meas.y(value) - 10, 20, 11, a.ref)


def _comments(canvas: Canvas, t: D.Template, L: dict) -> None:
    """Numbered margin comments, bold lead running straight into the body.

    IBCS puts explanation beside the chart rather than under it, tied to the mark
    by a circled number - so the reader can check the claim against the bar
    without moving their eyes off the page.
    """
    blue = S.ANNOTATION["highlight"]
    x, size = L["comments_x"], L["font"]["comment"]
    text_x = x + 18
    # Where the comments sit between panels rather than in the right margin,
    # the page edge is the wrong thing to wrap against - the next panel is.
    max_width = L.get("comments_width",
                      L["page"][0] - L["margin_right"] - text_x)

    for comment, y in zip(t.comments, L["comments_y"]):
        canvas.ellipse(x, y - 4, 9, 9, stroke=blue, stroke_width=1.2)
        canvas.text(x, y, comment.ref, size=size * 0.85, fill=blue, anchor="middle")
        if L.get("comment_lead_own_line"):
            # A lead on its own line where the column is narrow; run into the
            # body where it is wide. The reference does both, template by
            # template, and it is a wrapping decision rather than a rule.
            canvas.rich_text(text_x, y, [(comment.lead, "bold")], size=size)
            lines = _wrap_runs([(comment.body, "normal")], max_width, size=size)
            start = 1
        else:
            lines = _wrap_runs([(comment.lead, "bold"),
                                (" " + comment.body, "normal")],
                               max_width, size=size)
            start = 0
        for n, line in enumerate(lines):
            canvas.rich_text(text_x, y + (n + start) * 17, line, size=size)


# Approximate Arial advance widths, as a fraction of font size. Good enough to
# wrap fixed template copy; nothing here depends on the estimate being exact.
_NARROW = set("iljtfr.,;:'\"|!()[]{} ")
_WIDE = set("mwMW@")


def _text_width(text: str, size: float, bold: bool) -> float:
    total = 0.0
    for ch in text:
        if ch in _NARROW:
            total += 0.30
        elif ch in _WIDE:
            total += 0.85
        elif ch.isupper() or ch.isdigit():
            total += 0.62
        else:
            total += 0.53
    return total * size * (1.06 if bold else 1.0)


def _wrap_runs(runs: Sequence[tuple[str, str]], max_width: float,
               size: float) -> list[list[tuple[str, str]]]:
    """Wrap a sequence of (text, weight) runs to a pixel width, preserving weight."""
    words: list[tuple[str, str]] = []
    for content, weight in runs:
        for i, word in enumerate(content.split(" ")):
            if word:
                words.append((word, weight))
    lines: list[list[tuple[str, str]]] = []
    current: list[tuple[str, str]] = []
    width = 0.0
    for word, weight in words:
        w = _text_width(word + " ", size, weight == "bold")
        if current and width + w > max_width:
            lines.append(current)
            current, width = [], 0.0
        current.append((word + " ", weight))
        width += w
    if current:
        lines.append(current)
    return lines


# The attribution, in two forms. The first clause is a claim about provenance
# and it is only true of the seventeen recreations of the IBCS(R) Institute's
# own published examples - a chart drawn from someone's own figures is not a
# recreation of anything, so it takes the second string alone.
FOOTER_RECREATION = ("IBCS\u00ae template recreation \u00b7 Chart rendered with "
                     "Claude Code using a custom IBCS\u00ae skill")
FOOTER_OWN_DATA = ("Chart rendered with Claude Code using a custom "
                   "IBCS\u00ae skill")


def _footer(canvas: Canvas, L: dict) -> None:
    """One line, bottom left, over a hairline rule.

    Merged from a left and a right string, which read as two separate
    statements and left the eye no way to tell which was the attribution. The
    form is chosen by the layout rather than by the renderer, so a template
    recreation and a chart built from a reader's own data can share every line
    of drawing code and still say the right thing about themselves.
    """
    w, h = L["page"]
    canvas.line(L["margin_left"], h - 26, w - L["margin_right"], h - 26,
                stroke="#D9D9D9", stroke_width=1)
    canvas.text(L["margin_left"], h - 12,
                L.get("footer", FOOTER_RECREATION), size=9,
                fill=S.PAGE["footnote"])


# --------------------------------------------------------------------------- #
# C04A layout - measured from template-refs/C04_04A.png
# --------------------------------------------------------------------------- #
#
# C04A is C03A transposed and re-referenced: bars not columns, plan not prior
# year, and the three tiers sit side by side sharing one vertical category axis
# rather than stacked sharing a horizontal one.
#
# Three vertical offsets, all measured rather than assumed, and all different:
# the PL bar sits 4px above the row centre, the AC bar and the dPL bar sit on
# it, and the dPL% pin sits 2px above it. The variance bar tracking the AC bar
# rather than the row centre is the sensible part - the variance belongs to the
# actual - and the pin's 2px is the reference render's own inconsistency,
# reproduced because this is a recreation.

C04A_LAYOUT = dict(
    page=(1280, 720),
    margin_left=20,
    margin_right=20,
    subject_suffix=" (sorted by ΔPL)",
    message_width=430,        # breaks after "(+31 kUSD)", as the original does

    row0=130.0,               # centre of the first category row
    cat_pitch=25.611,         # (591 - 130) / 18
    bar_height=19,
    pl_offset=-4,             # PL sits above the AC bar, as PY sits left in C03A
    pin_offset=-2,

    label_right=172,          # category names are right-aligned to here
    axis_top=104,
    axis_bottom=603.8,        # row0 - pitch/2 + pitch * 19

    # The measure and dPL panels share one scale, because they share a unit.
    # check_scale() in ibcs_layout enforces the same rule on the Excel side.
    measure_zero=179.5,
    unit_scale=1.603,         # px per kUSD
    abs_zero=787.0,
    rel_zero=1011.0,
    rel_scale=0.80,           # px per percentage point - a different unit, so free
    pin_stem=5,
    pin_head=5,

    # Reading rules every five rows. Not decoration: nineteen unbroken rows of
    # bars is hard to track across three panels.
    group_after=(3, 8, 13),
    group_rule_offset=-2,

    total_rule_y=606.5,
    total_row=619.0,
    total_legend_x=(232.0, 313.0),

    header_y=110,
    title_rule_y=88,
    font=dict(entity=13, measure=13.5, period=13, message=13.5, badge=34,
              value=13.5, category=13.5, header=13, total=13.5, footer=9),
)


def render_c04a(template: D.Template = D.C04A, layout: dict | None = None) -> str:
    """Render C04A - Housing and Construction net sales by state against plan."""
    L = {**C04A_LAYOUT, **(layout or {})}
    t = template
    width, height = L["page"]
    c = Canvas(width, height, prefix=f"{t.id}{t.variant}".lower())
    rows = _C04Rows(L, len(t.categories))

    _title_block(c, t, L)
    _c04_group_rules(c, L, rows)
    _c04_measure_panel(c, t, L, rows)
    _c04_abs_panel(c, t, L, rows)
    _c04_rel_panel(c, t, L, rows)
    _c04_category_labels(c, t, L, rows)
    _c04_totals(c, t, L)
    _c04_annotations(c, t, L, rows)
    _footer(c, L)
    return c.to_svg()


class _C04Rows:
    """The shared category axis, as three offsets off one row centre.

    Every panel asks this for its y, so the panels cannot drift relative to one
    another - the same guarantee the CategoryAxis gives C03A, and the reason the
    Excel port has to work so hard to reproduce it.
    """

    def __init__(self, L: dict, count: int) -> None:
        self.L, self.count = L, count

    def centre(self, i: int) -> float:
        return self.L["row0"] + self.L["cat_pitch"] * i

    def pl(self, i: int) -> float:
        return self.centre(i) + self.L["pl_offset"]

    def pin(self, i: int) -> float:
        return self.centre(i) + self.L["pin_offset"]

    def boundary(self, i: int) -> float:
        """The line between row i and row i+1."""
        return self.centre(i) + self.L["cat_pitch"] / 2


def _c04_measure_panel(c: Canvas, t: D.Template, L: dict, rows: _C04Rows) -> None:
    """PL outlined behind, AC solid in front, both labelled where they fit."""
    tier = t.tier("measure")
    # The reference scenario comes from the variance tier that names it, not
    # from a constant. Plan and budget are notated identically - both are
    # fictitious, so both are outlined - and which one a chart is measured
    # against is a fact about the data, not about the drawing.
    reference = t.tier("var_abs").reference or "PL"
    pl, ac = tier.series_for(reference), tier.series_for("AC")
    zero, scale, F = L["measure_zero"], L["unit_scale"], L["font"]

    c.line(zero, L["axis_top"], zero, L["axis_bottom"],
           stroke=S.TEXT["primary"], stroke_width=2)

    for i in range(len(t.categories)):
        pl_x = zero + pl.values[i] * scale
        ac_x = zero + ac.values[i] * scale
        scenario_bar_h(c, reference, rows.pl(i), L["bar_height"], pl_x, zero)
        scenario_bar_h(c, "AC", rows.centre(i), L["bar_height"], ac_x, zero)

        ac_text = tier.number_format.format(ac.values[i])
        c.text(ac_x + 5, rows.centre(i) + 4, ac_text, size=F["value"])

        # The plan bar is only labelled where its label would not collide with
        # the actual's - which is why the reference shows 88 against Missouri
        # but nothing against Wisconsin, whose plan is only 6 kUSD away.
        gap = pl_x - ac_x
        if gap > _text_width(ac_text, F["value"], False) + 12:
            c.text(pl_x + 5, rows.centre(i) + 4,
                   tier.number_format.format(pl.values[i]), size=F["value"])

    # Integrated legend: name each series where it is, not in a box elsewhere.
    c.text((zero + zero + pl.values[0] * scale) / 2, L["header_y"], reference,
           size=F["header"], anchor="middle")
    c.text(zero + ac.values[0] * scale + 6, L["header_y"], "AC", size=F["header"])


def _c04_abs_panel(c: Canvas, t: D.Template, L: dict, rows: _C04Rows) -> None:
    """Absolute variances: full-height bars coloured by impact, on a double rule."""
    tier = t.tier("var_abs")
    values = tier.series_for("AC").values
    zero, scale, F = L["abs_zero"], L["unit_scale"], L["font"]

    reference_axis_v(c, tier.reference, L["axis_top"], L["axis_bottom"], zero)

    for i, value in enumerate(values):
        x = zero + value * scale
        colour = S.variance_colour(value, tier.higher_is_better)
        left, right = sorted((x, zero))
        c.rect(left, rows.centre(i) - L["bar_height"] / 2,
               right - left, L["bar_height"], fill=colour)
        _c04_value_label(c, tier, i, value, x, rows.centre(i) + 4, F["value"])

    c.text(zero + 10, L["header_y"], tier.label, size=F["header"])


def _c04_rel_panel(c: Canvas, t: D.Template, L: dict, rows: _C04Rows) -> None:
    """Relative variances as pins - a thin stem with a head carrying AC's fill."""
    tier = t.tier("var_rel")
    values = tier.series_for("AC").values
    zero, scale, F = L["rel_zero"], L["rel_scale"], L["font"]

    reference_axis_v(c, tier.reference, L["axis_top"], L["axis_bottom"], zero)

    for i, value in enumerate(values):
        x = zero + value * scale
        colour = S.variance_colour(value, tier.higher_is_better)
        y = rows.pin(i)
        left, right = sorted((x, zero))
        c.rect(left, y - L["pin_stem"] / 2, right - left, L["pin_stem"], fill=colour)

        head = L["pin_head"]
        fill, outline = paint_for(c, t.category_scenarios[i])
        c.rect(x - head / 2, y - head / 2, head, head,
               fill=fill, stroke=outline or fill, stroke_width=0.8)
        _c04_value_label(c, tier, i, value, x, y + 4, F["value"], pad=6)

    c.text(zero + 2, L["header_y"], tier.label, size=F["header"])


def _c04_value_label(c: Canvas, tier: D.Tier, index: int, value: float,
                     x: float, y: float, size: float, pad: float = 5) -> None:
    """Outside the element, in the direction it points - never over the axis."""
    text = tier.label_for(index, value)
    if value >= 0:
        c.text(x + pad, y, text, size=size)
    else:
        c.text(x - pad, y, text, size=size, anchor="end")


def _c04_category_labels(c: Canvas, t: D.Template, L: dict, rows: _C04Rows) -> None:
    for i, name in enumerate(t.categories):
        c.text(L["label_right"], rows.centre(i) + 4, name,
               size=L["font"]["category"], anchor="end")


def _c04_group_rules(c: Canvas, L: dict, rows: _C04Rows) -> None:
    """Light rules every five rows, so the eye can track one state across panels."""
    for i in L["group_after"]:
        y = rows.boundary(i) + L["group_rule_offset"]
        c.line(L["label_right"] + 6, y, L["rel_zero"] + 120, y,
               stroke=S.PAGE["rule"], stroke_width=1)


def _c04_totals(c: Canvas, t: D.Template, L: dict) -> None:
    """The USA row: the two totals as an integrated legend, plus their variances.

    Every number here is the one IBCS printed, taken from the template's
    summary rather than recomputed. The transcribed actuals add to 1 844 against
    a printed 1 845, because nineteen rounded values need not add to the rounded
    total - and dPL% cannot be summed at all.
    """
    F, y = L["font"], L["total_row"]
    summary = t.summary
    # A template with nothing to total draws no total row. C04A's own data has
    # one, but a chart of nineteen cost lines that do not add to anything a
    # reader wants stated should not invent a rule and a blank label.
    if summary is None:
        return

    c.line(L["label_right"] + 6, L["total_rule_y"], L["rel_zero"] + 120,
           L["total_rule_y"], stroke=S.TEXT["primary"], stroke_width=2)
    c.text(L["label_right"], y + 4, summary.label, size=F["category"],
           weight="bold", anchor="end")
    c.line(L["measure_zero"], L["total_rule_y"] + 4, L["measure_zero"], y + 12,
           stroke=S.TEXT["primary"], stroke_width=2)

    # The totals are given as a legend rather than as bars: they are an
    # aggregate of the rows above, not a twentieth state.
    for x, (scenario, total) in zip(L["total_legend_x"], summary.stack):
        fill, outline = paint_for(c, scenario)
        c.rect(x, y - 8, 10, 10, fill=fill, stroke=outline or fill, stroke_width=1)
        c.text(x + 16, y + 4, _thousands(total), size=F["total"], weight="bold")

    for key, zero, scale, pad, total in (
        ("var_abs", L["abs_zero"], L["unit_scale"], 5, summary.variance_abs),
        ("var_rel", L["rel_zero"], L["rel_scale"], 6, summary.variance_rel),
    ):
        tier = t.tier(key)
        x = zero + total * scale
        colour = S.variance_colour(total, tier.higher_is_better)
        height = L["bar_height"] if key == "var_abs" else L["pin_stem"]
        left, right = sorted((x, zero))
        c.rect(left, y - height / 2 - 4, right - left, height, fill=colour)
        text = tier.number_format.format(total)
        if total >= 0:
            c.text(x + pad, y + 4, text, size=F["value"], weight="bold")
        else:
            c.text(x - pad, y + 4, text, size=F["value"], weight="bold",
                   anchor="end")


def _c04_annotations(c: Canvas, t: D.Template, L: dict, rows: _C04Rows) -> None:
    """Highlight ovals round the values the message names."""
    for note in t.annotations:
        key, index = note.target
        if key != "var_abs":
            continue
        tier = t.tier(key)
        value = tier.series_for("AC").values[index]
        x = L["abs_zero"] + value * L["unit_scale"]
        text = tier.label_for(index, value)
        w = _text_width(text, L["font"]["value"], False)
        cx = x + 5 + w / 2
        cy = rows.centre(index) + 1
        if note.kind == "oval":
            highlight_oval(c, cx, cy, w / 2 + 9, 10, note.ref)
        elif note.kind == "arrow":
            blue = S.ANNOTATION["highlight"]
            top = cy - 34
            c.line(cx, top, cx, cy - 12, stroke=blue, stroke_width=3)
            c.add(f'<path d="M {cx - 6:.2f} {cy - 16:.2f} L {cx + 6:.2f} '
                  f'{cy - 16:.2f} L {cx:.2f} {cy - 8:.2f} Z" fill="{blue}"/>')


# --------------------------------------------------------------------------- #
# C12A layout - measured from template-refs/C12_12A-1.png
# --------------------------------------------------------------------------- #
#
# Four panels sharing one vertical category axis: two waterfalls, then the
# absolute and relative variances between them. Three of the four are in kEUR
# and all three are drawn at 0.3085 px per kEUR - measured, not assumed, and the
# thing that makes the dPY column readable as a piece of the AC panel rather
# than as its own chart.
#
# The row pitch is not uniform. A result subtotal is followed by a wider gap
# than an ordinary line, which is what separates the statement into its blocks,
# so row positions are computed from a count of the subtotals above rather than
# from the index. Fitting the measured centres needs both numbers: 25.6 and an
# extra 6.4 after each of the five result lines, which reproduces all twenty
# rows to within a pixel.

C12A_LAYOUT = dict(
    page=(1280, 720),
    margin_left=20,
    margin_right=20,
    message_width=600,
    message_x=401,

    row0=122.0,               # centre of the first row
    row_pitch=25.6,
    row_extra=6.4,            # added after each result subtotal
    bar_height=19,

    label_x=18,               # the prefix column
    label_indent=25,          # components of a subtotal sit in from it

    # Three panels, one unit, one scale.
    py_zero=205.5,
    ac_zero=509.5,
    abs_zero=919.5,
    unit_scale=0.3085,        # px per kEUR

    # The relative panel is a percentage, so its scale is free.
    rel_zero=1104.0,
    rel_scale=0.60,           # px per percentage point
    rel_limit=1265.0,         # where the panel runs out and outliers begin
    pin_stem=5,
    pin_head=5,

    rule_x=(150.0, 1265.0),   # the block rules, above each result subtotal
    rule_weight=2,
    axis_weight=2,           # the waterfall's own zero line
    connector_weight=2,
    connector_colour="#7F7F7F",
    rule_lift=13.5,           # above the row centre

    title_rule_y=73.5,
    header_y=101,
    header_x=dict(wf_py=317, wf_ac=649, var_abs=949, var_rel=1112),

    outlier_size=9,           # triangle edge
    outlier_gap=3,

    font=dict(entity=13, measure=13.5, period=13, message=13.5, badge=34,
              value=13.5, category=13.5, header=13, total=13.5, footer=9),
)


class _C12Rows:
    """The shared category axis of a statement.

    Unlike a month or a state, the rows of a statement are not evenly spaced:
    the blocks a reader thinks in - revenue, operating result, result after tax
    - are separated by extra space and a rule. Every panel asks this object for
    its y, so the four panels cannot drift apart, and the block spacing is
    derived from the rows themselves rather than from a list of pixel offsets
    that would have to be rewritten if a line were inserted.
    """

    def __init__(self, L: dict, rows) -> None:
        self.L, self.rows = L, rows

    def _breaks_before(self, index: int) -> int:
        return sum(1 for r in self.rows[:index] if r.spans == "zero")

    def centre(self, i: int) -> float:
        return (self.L["row0"] + self.L["row_pitch"] * i
                + self.L["row_extra"] * self._breaks_before(i))

    def top(self) -> float:
        return self.centre(0) - self.L["row_pitch"] / 2

    def bottom(self) -> float:
        return self.centre(len(self.rows) - 1) + self.L["row_pitch"] / 2


def outlier_marker(canvas: Canvas, x: float, cy: float, colour: str,
                   count: int, direction: int, L: dict) -> None:
    """Triangles saying the value continues past the edge of the panel.

    Rule UN 5.3: where a relative variance is enormous because its base is tiny,
    IBCS marks it rather than rescaling every other row to it - and the marker
    is specified as "omit the pin head and add outlier triangles pointing in the
    direction of growth". The head is omitted because a head states a measured
    value at a position, and this position is not where the value is.

    The *number* of triangles is not in the standard. C12A draws three for
    +983% and one for +391%, which is one per panel-width the value overruns
    by, and that is what is reproduced here - as an observed house choice, not
    as a rule.
    """
    size, gap = L["outlier_size"], L["outlier_gap"]
    for n in range(count):
        tip = x + direction * n * (size * 0.62 + gap)
        back = tip - direction * size * 0.62
        canvas.add(
            f'<polygon points="{tip:.2f},{cy:.2f} {back:.2f},{cy - size / 2:.2f} '
            f'{back:.2f},{cy + size / 2:.2f}" fill="{colour}"/>'
        )


def halo_text(canvas: Canvas, x: float, y: float, content: str, size: float,
              anchor: str = "start") -> None:
    """A value label that has to stay readable where it sits over its own bar.

    Only the outlier rows need this - their bars run the width of the panel and
    the label has nowhere clear to go - and the original solves it the same way,
    with the text knocked out of the bar rather than moved off it.
    """
    canvas.add(
        f'<text x="{x:.2f}" y="{y:.2f}" font-family="Arial, Helvetica, sans-serif" '
        f'font-size="{size}" text-anchor="{anchor}" fill="#FFFFFF" '
        f'stroke="#FFFFFF" stroke-width="3.5" stroke-linejoin="round" '
        f'>{content}</text>'
    )
    canvas.text(x, y, content, size=size, anchor=anchor)


def render_c12a(template: D.Template = D.C12A, layout: dict | None = None) -> str:
    """Render C12A - a profit and loss statement as two waterfalls and a variance."""
    L = {**C12A_LAYOUT, **(layout or {})}
    t = template
    width, height = L["page"]
    c = Canvas(width, height, prefix=f"{t.id}{t.variant}".lower())
    rows = _C12Rows(L, t.rows)

    _title_block(c, t, L)
    _c12_block_rules(c, t, L, rows)
    # The scenario each panel draws comes from the tier, not from a constant.
    # A statement compared against budget rather than prior year is the same
    # chart with a different notation on its left-hand panel - outlined because
    # a budget is a plan, where prior year is solid because it happened - and
    # hard-coding "PY" made that chart impossible to draw.
    _c12_waterfall(c, t, L, rows, "wf_py", None, L["py_zero"])
    _c12_waterfall(c, t, L, rows, "wf_ac", None, L["ac_zero"])
    _c12_abs_panel(c, t, L, rows)
    _c12_rel_panel(c, t, L, rows)
    _c12_row_labels(c, t, L, rows)
    _c12_headers(c, t, L)
    _c12_annotations(c, t, L, rows)
    _footer(c, L)
    return c.to_svg()


def _c12_block_rules(c: Canvas, t: D.Template, L: dict, rows: _C12Rows) -> None:
    """A rule above each result line, which is what makes the blocks visible."""
    x0, x1 = L["rule_x"]
    for i, row in enumerate(t.rows):
        if row.spans == "zero":
            y = rows.centre(i) - L["rule_lift"]
            c.line(x0, y, x1, y, stroke=S.TEXT["primary"],
                   stroke_width=L["rule_weight"])


def _c12_connectors(c: Canvas, t: D.Template, L: dict, rows: _C12Rows,
                    spans, zero: float) -> None:
    """The lines that make a waterfall a waterfall rather than twenty bars.

    Each one marks the level the statement has reached, carried from the bar
    that ended there to the bar that starts there. Without them the reader has
    to infer that a bar begins where the last one stopped, which is the single
    thing the chart form exists to show.

    Two kinds, and the second is easy to miss. Every adjacent pair gets a
    connector at the running level between them. A *group* subtotal gets one
    more: a line at its far end, running back up to the row where its block
    began, which is what says the bar is a bracket over those lines rather than
    another step in the sequence.
    """
    scale = L["unit_scale"]
    colour, weight = L["connector_colour"], L["connector_weight"]

    for i in range(len(t.rows) - 1):
        x = zero + spans[i][1] * scale
        c.line(x, rows.centre(i), x, rows.centre(i + 1),
               stroke=colour, stroke_width=weight)

    for i, row in enumerate(t.rows):
        if row.kind == "subtotal" and row.spans != "zero":
            x = zero + spans[i][0] * scale
            c.line(x, rows.centre(i - int(row.spans) - 1), x, rows.centre(i),
                   stroke=colour, stroke_width=weight)


def _c12_waterfall(c: Canvas, t: D.Template, L: dict, rows: _C12Rows,
                   key: str, scenario: str, zero: float) -> None:
    """One scenario's statement, walked.

    Where each bar starts and ends comes from ``waterfall_spans`` in the data
    layer rather than from arithmetic here, so the Excel build and this drawing
    cannot disagree about the running total - which they would, because the rule
    has three cases and only the element case is obvious.
    """
    tier = t.tier(key)
    # None means "whatever this tier carries", which is the normal case; a
    # named scenario still works for a caller that wants to pick one out of a
    # tier holding several.
    series = tier.series_for(scenario) if scenario else tier.series[0]
    scenario = series.scenario
    values = series.values
    scale, F = L["unit_scale"], L["font"]
    spans = D.waterfall_spans(t, values)

    # Connectors first: bars are opaque and must cover the ends of the lines
    # they meet, not be crossed by them.
    _c12_connectors(c, t, L, rows, spans, zero)
    c.line(zero, rows.top(), zero, rows.bottom(),
           stroke=S.TEXT["primary"], stroke_width=L["axis_weight"])

    for i, (row, (lo, hi)) in enumerate(zip(t.rows, spans)):
        x0, x1 = sorted((zero + lo * scale, zero + hi * scale))
        c.rect(x0, rows.centre(i) - L["bar_height"] / 2, max(x1 - x0, 1.0),
               L["bar_height"], fill=S.waterfall_fill(scenario, row.sign))

        # The label goes on the outside of the bar, on the side the statement
        # reads towards - which for a deduction is the left.
        text = _thousands(values[i])
        bold = row.kind == "subtotal"
        if row.sign > 0:
            c.text(x1 + 5, rows.centre(i) + 4, text, size=F["value"], weight="bold" if bold else "normal")
        else:
            c.text(x0 - 5, rows.centre(i) + 4, text, size=F["value"],
                   anchor="end", weight="bold" if bold else "normal")


def _c12_abs_panel(c: Canvas, t: D.Template, L: dict, rows: _C12Rows) -> None:
    """Absolute variances, coloured by the impact each *line* has."""
    tier = t.tier("var_abs")
    values = tier.series_for("AC").values
    zero, scale, F = L["abs_zero"], L["unit_scale"], L["font"]

    reference_axis_v(c, tier.reference, rows.top(), rows.bottom(), zero)

    for i, value in enumerate(values):
        # Not tier.higher_is_better: on a statement the direction is a property
        # of the line. Revenue up is good, cost up is not, in the same column.
        good = t.higher_is_better_at(i, tier)
        colour = S.variance_colour(value, good)
        x = zero + value * scale
        left, right = sorted((x, zero))
        # Nothing drawn for a variance of nothing - see _c12_rel_panel.
        if value:
            c.rect(left, rows.centre(i) - L["bar_height"] / 2,
                   max(right - left, 1.0), L["bar_height"], fill=colour)

        text = tier.label_for(i, value)
        bold = t.rows[i].kind == "subtotal"
        if value >= 0:
            c.text(x + 6, rows.centre(i) + 4, text, size=F["value"], weight="bold" if bold else "normal")
        else:
            c.text(x - 6, rows.centre(i) + 4, text, size=F["value"],
                   anchor="end", weight="bold" if bold else "normal")


def _c12_rel_panel(c: Canvas, t: D.Template, L: dict, rows: _C12Rows) -> None:
    """Relative variances as pins, with outlier markers where they run off."""
    tier = t.tier("var_rel")
    values = tier.series_for("AC").values
    zero, scale, F = L["rel_zero"], L["rel_scale"], L["font"]
    limit = L["rel_limit"]
    head = L["pin_head"]

    reference_axis_v(c, tier.reference, rows.top(), rows.bottom(), zero)

    for i, value in enumerate(values):
        good = t.higher_is_better_at(i, tier)
        colour = S.variance_colour(value, good)
        cy = rows.centre(i)
        x = zero + value * scale
        text = tier.label_for(i, value)
        bold = t.rows[i].kind == "subtotal"

        # Derived, not listed. This used to test the row's *index* against a
        # set transcribed from IBCS's own statement, which is fine until the
        # chart is asked to draw somebody else's - then row 3 of a completely
        # different P&L gets clipped and given arrows because row 3 of theirs
        # was. Asking whether the bar actually runs off the panel reproduces
        # their two outliers exactly and works for any data.
        if abs(value) * scale > limit - zero:
            # Clipped to the panel, no head, triangles instead. How far it
            # overruns decides how many, which is what the original does.
            stop = limit - L["outlier_size"] * 2.4
            c.rect(zero, cy - L["pin_stem"] / 2, stop - zero, L["pin_stem"],
                   fill=colour)
            count = max(1, min(3, int(abs(value) / ((limit - zero) / scale))))
            outlier_marker(c, stop + 2, cy, colour, count, +1, L)
            halo_text(c, stop - 2, cy + 4, text, F["value"], anchor="end")
            continue

        left, right = sorted((x, zero))
        # A variance of exactly nothing draws nothing. The minimum length below
        # exists so a small variance still reads as a mark, but applied to zero
        # it puts a stub on the axis in the neutral colour - which says "no
        # change" in a notation where every other mark says how much.
        if value:
            c.rect(left, cy - L["pin_stem"] / 2, max(right - left, 1.0),
                   L["pin_stem"], fill=colour)
        # The head carries the minuend's scenario fill - AC, since the variance
        # is AC against PY - which is what says the value was measured.
        fill, outline = paint_for(c, "AC")
        hx = x + (0 if value >= 0 else -head)
        c.rect(hx, cy - head / 2, head, head, fill=fill,
               stroke=outline or fill, stroke_width=0.8)
        if value >= 0:
            c.text(hx + head + 5, cy + 4, text, size=F["value"], weight="bold" if bold else "normal")
        else:
            c.text(hx - 5, cy + 4, text, size=F["value"], anchor="end", weight="bold" if bold else "normal")


def _c12_row_labels(c: Canvas, t: D.Template, L: dict, rows: _C12Rows) -> None:
    """The statement itself, down the left: prefix, indent and weight."""
    F = L["font"]
    for i, row in enumerate(t.rows):
        y = rows.centre(i) + 4
        bold = row.kind == "subtotal"
        x = L["label_x"] + (L["label_indent"] if row.indent else 0)
        if row.prefix:
            c.text(L["label_x"], y, row.prefix, size=F["category"], weight="bold" if bold else "normal")
            x = L["label_x"] + 14
        c.text(x, y, row.label, size=F["category"], weight="bold" if bold else "normal")


def _c12_headers(c: Canvas, t: D.Template, L: dict) -> None:
    """One header per panel. This is the integrated legend for a waterfall.

    The two waterfall panels are the same measure in two scenarios, and nothing
    inside either panel says which is which - the greys separate adding from
    subtracting, not actual from prior year. So the header is load-bearing here
    in a way a column chart's is not.
    """
    F = L["font"]
    for tier in t.tiers:
        x = L["header_x"].get(tier.key)
        if x is not None:
            c.text(x, L["header_y"], tier.label, size=F["header"], anchor="middle")


def _c12_annotations(c: Canvas, t: D.Template, L: dict, rows: _C12Rows) -> None:
    """Ovals round the three numbers the message names."""
    tier = t.tier("var_abs")
    values = tier.series_for("AC").values
    zero, scale = L["abs_zero"], L["unit_scale"]
    for note in t.annotations:
        if note.kind != "oval":
            continue
        key, i = note.target
        if key != "var_abs":
            continue
        x = zero + values[i] * scale
        text = tier.label_for(i, values[i])
        half = _text_width(text, L["font"]["value"], True) / 2 + 7
        cx = x + 6 + half - 7 if values[i] >= 0 else x - 6 - half + 7
        highlight_oval(c, cx, rows.centre(i), half, 11)


# --------------------------------------------------------------------------- #
# C06F layout - measured from template-refs/C06_06F.png
# --------------------------------------------------------------------------- #
#
# Three scenario bars, fifteen state bars between them, a waterfall bridging the
# PY bar to the AC bar, and a pin panel. The bridge is what makes this template
# different from C12A: the waterfall does not walk up from zero, it starts on
# one total and lands on another, and the two totals are drawn as bars so the
# reader can see it land.
#
# One scale for the whole left side, measured from the two totals themselves:
# AC 1728 ends at x 771 and PY 2071 at x 897, which fixes 0.3673 px per kUSD and
# an origin at 136.3. Every state bar and every waterfall step is drawn on it, so
# a step is directly comparable with the bar it came from.

C06F_LAYOUT = dict(
    page=(1280, 720),
    margin_left=20,
    margin_right=20,
    subject_suffix=" (sorted by ΔPY)",
    message_width=560,
    title_rule_y=76.5,

    zero_x=136.3,
    unit_scale=0.3673,            # px per kUSD, shared by the bars and the bridge

    row_pl=131.0,
    row_py=163.0,
    row0=197.0,                   # first state
    row_pitch=24.0,
    row_ac=571.0,
    bar_height=22,                # a state bar
    scenario_height=18,           # PL, PY and AC are drawn slightly thinner
    step_height=18,               # a waterfall step
    total_height=8,               # the two variance bars under the waterfall

    wf_total_y=616.5,
    pl_var_y=660.5,

    label_right=133.0,            # state names are right-aligned to here
    value_gap=6.0,

    # A rule every four states. Fifteen unbroken rows across four panels is more
    # than a reader can track, and the reference breaks them into blocks of four.
    rule_every=4,
    rule_x=(139.0, 1259.0),
    rule_lift=10.5,

    rel_zero=1143.0,
    rel_scale=2.6,                # px per percentage point - its own unit, so free
    rel_header_x=1149.5,
    header_y=170.0,

    pin_stem=5,
    pin_head=5,

    font=dict(entity=13, measure=13.5, period=13, message=13.5, badge=34,
              value=13, category=13, header=13, total=13.5, footer=9),
)


class _C06Rows:
    """The shared category axis: three scenario rows and fifteen states.

    The scenario rows are not categories - they are totals of the categories -
    so they sit outside the pitch rather than being row -1 and row 15. Keeping
    them out of the arithmetic is what lets the pitch stay uniform and the
    reading rules fall on multiples of four.
    """

    def __init__(self, L: dict, count: int) -> None:
        self.L, self.count = L, count

    def centre(self, i: int) -> float:
        return self.L["row0"] + self.L["row_pitch"] * i

    def top(self) -> float:
        return self.centre(0) - self.L["row_pitch"] / 2

    def bottom(self) -> float:
        return self.centre(self.count - 1) + self.L["row_pitch"] / 2


def render_c06f(template: D.Template = D.C06F, layout: dict | None = None) -> str:
    """Render C06F - net sales by state, bars bridged to prior year by a waterfall."""
    L = {**C06F_LAYOUT, **(layout or {})}
    t = template
    width, height = L["page"]
    c = Canvas(width, height, prefix=f"{t.id}{t.variant}".lower())
    rows = _C06Rows(L, len(t.categories))

    _title_block(c, t, L)
    _c06_rules(c, t, L, rows)
    _c06_guides(c, t, L)
    _c06_scenario_bars(c, t, L)
    _c06_state_bars(c, t, L, rows)
    _c06_waterfall(c, t, L, rows)
    _c06_totals(c, t, L)
    _c06_pins(c, t, L, rows)
    _c06_labels(c, t, L, rows)
    _c06_annotations(c, t, L, rows)
    _footer(c, L)
    return c.to_svg()


def _c06_x(L: dict, value: float) -> float:
    return L["zero_x"] + value * L["unit_scale"]


def _c06_rules(c: Canvas, t: D.Template, L: dict, rows: _C06Rows) -> None:
    x0, x1 = L["rule_x"]
    for i in range(L["rule_every"] - 1, len(t.categories) - 1, L["rule_every"]):
        y = rows.centre(i) + L["rule_lift"]
        c.line(x0, y, x1, y, stroke="#BFBFBF", stroke_width=1)


def _c06_guides(c: Canvas, t: D.Template, L: dict) -> None:
    """Vertical lines at the plan, actual and prior-year totals.

    These are what make the waterfall legible as a bridge: the step column sits
    between the PY line and the AC line, so where it starts and where it has to
    finish are both visible while reading it.
    """
    for label, y0, y1 in (("PL", L["row_pl"] - 9, L["pl_var_y"] + 6),
                          ("AC", L["row_pl"] - 9, L["pl_var_y"] + 22),
                          ("PY", L["row_py"] - 9, L["wf_total_y"] + 6)):
        value = t.summary_row(label).total
        x = _c06_x(L, value)
        c.line(x, y0, x, y1, stroke=S.TEXT["primary"], stroke_width=1)


def _c06_scenario_bars(c: Canvas, t: D.Template, L: dict) -> None:
    """PL outlined, PY solid light, AC solid dark - the notation, unmixed.

    Three scenarios in one column is the clearest statement of UN 3.2 anywhere
    in the template library: the same measure, drawn three ways, and the fill is
    the only thing saying which is which.
    """
    F, zero = L["font"], L["zero_x"]
    for scenario, y in (("PL", L["row_pl"]), ("PY", L["row_py"]),
                        ("AC", L["row_ac"])):
        value = t.summary_row(scenario).total
        x = _c06_x(L, value)
        scenario_bar_h(c, scenario, y, L["scenario_height"], x, zero)
        c.text(L["label_right"], y + 4, scenario, size=F["category"],
               anchor="end", weight="bold")
        c.text(x + L["value_gap"], y + 4, _thousands(value), size=F["total"],
               weight="bold")


def _c06_state_bars(c: Canvas, t: D.Template, L: dict, rows: _C06Rows) -> None:
    """AC solid, and where prior year was higher, the shortfall appended in its
    own fill.

    So the bar reaches PY and the grey part is exactly what was lost - which is
    why Illinois reads at a glance and is the state the message is about. Only a
    shortfall wide enough to hold its own number is labelled; the reference
    labels Illinois and nothing else, and the same rule reproduces that.
    """
    tier = t.tier("measure")
    ac = tier.series_for("AC").values
    py = tier.series_for("PY").values
    F, zero = L["font"], L["zero_x"]

    # The zero axis the bars stand on. Two pixels of black, and easy to leave
    # out because the bars start there anyway - until a row's value is small
    # enough that there is nothing to mark where the column begins.
    c.line(zero, L["row_pl"] - 11, zero, L["row_ac"] + 11,
           stroke=S.TEXT["primary"], stroke_width=2)

    # Bars first, all of them, then the labels. Drawing each row's label before
    # the next element meant the shortfall was painted over the number it
    # belonged to - visible only on the rows where a shortfall existed, which
    # made it look like a font problem rather than an ordering one.
    for i in range(len(t.categories)):
        y = rows.centre(i)
        scenario_bar_h(c, "AC", y, L["bar_height"], _c06_x(L, ac[i]), zero)
        if py[i] > ac[i]:
            scenario_bar_h(c, "PY", y, L["bar_height"], _c06_x(L, py[i]),
                           _c06_x(L, ac[i]))

    for i in range(len(t.categories)):
        y = rows.centre(i)
        ac_x, py_x = _c06_x(L, ac[i]), _c06_x(L, py[i])
        ac_text = f"{ac[i]:,.0f}"
        shortfall = py_x - ac_x if py[i] > ac[i] else 0.0

        # The actual's number always sits against its own bar, even where the
        # shortfall runs on past it - measured off the reference, which prints
        # California's 257 over the grey rather than clear of it. The number
        # belongs to the dark bar, so it goes where the dark bar ends.
        c.text(ac_x + L["value_gap"], y + 4, ac_text, size=F["value"])

        # The prior year is labelled only where its own segment has room, which
        # in this variant means Illinois and nothing else. Decided by measuring
        # the segment rather than by naming the state, so it stays right if a
        # figure is retyped.
        if shortfall > _text_width(ac_text, F["value"], False) + 24:
            c.text(py_x + L["value_gap"], y + 4, f"{py[i]:,.0f}", size=F["value"])


def _c06_waterfall(c: Canvas, t: D.Template, L: dict, rows: _C06Rows) -> None:
    """The bridge: one step per state, carrying the PY total onto the AC total.

    The walk comes from the data layer, started at the prior-year total rather
    than at zero - which is the whole difference between a bridge and a
    statement, and the reason ``waterfall_spans`` takes a starting level.
    """
    tier = t.tier("wf")
    values = tier.series_for("AC").values
    spans = D.waterfall_spans(t, values, start=t.summary_row("PY").total)
    F = L["font"]

    for i, (lo, hi) in enumerate(spans):
        y = rows.centre(i)
        x0, x1 = sorted((_c06_x(L, lo), _c06_x(L, hi)))
        colour = S.variance_colour(values[i], tier.higher_is_better)
        # A floor of three pixels. Four of these states moved by 1 to 3 kUSD,
        # which is under a pixel at this scale, and a step that rounds away
        # leaves a gap in the bridge - the reader sees the walk stop. Three
        # rather than one because a sub-pixel rectangle renders as two
        # half-covered pixels and reads as a smudge; the reference draws three.
        c.rect(x0, y - L["step_height"] / 2, max(x1 - x0, 3.0), L["step_height"],
               fill=colour)

        text = tier.label_for(i, values[i])
        if values[i] >= 0:
            c.text(x1 + L["value_gap"], y + 4, text, size=F["value"])
        else:
            c.text(x0 - L["value_gap"], y + 4, text, size=F["value"], anchor="end")


def _c06_totals(c: Canvas, t: D.Template, L: dict) -> None:
    """The two variance bars under the bridge: AC against PY, and AC against PL.

    Both are drawn between the guides they compare, so neither needs an axis -
    the bar starts on one total's line and ends on the other's.
    """
    F = L["font"]
    # Each bar runs between the two totals it compares, and which two those are
    # is stated by the summary row rather than by this function - so a variance
    # bar cannot end up drawn against a total it is not a variance from.
    for row, y, label in ((t.summary_row("ΔPY"), L["wf_total_y"], False),
                          (t.summary_row("vs PL"), L["pl_var_y"], True)):
        lo, hi = (t.summary_row(name).total for name in row.span)
        x0, x1 = sorted((_c06_x(L, lo), _c06_x(L, hi)))
        c.rect(x0, y - L["total_height"] / 2, x1 - x0, L["total_height"],
               fill=S.variance_colour(row.variance_abs, True))
        if label:
            c.text((x0 + x1) / 2, y + 20, f"{row.variance_abs:+.0f}",
                   size=F["value"], anchor="middle")


def _c06_pins(c: Canvas, t: D.Template, L: dict, rows: _C06Rows) -> None:
    """Relative variances as pins, with the summary pin on the AC row."""
    tier = t.tier("var_rel")
    values = tier.series_for("AC").values
    zero, scale, F = L["rel_zero"], L["rel_scale"], L["font"]
    head = L["pin_head"]

    reference_axis_v(c, tier.reference, rows.top(), rows.bottom(), zero)

    entries = [(rows.centre(i), values[i], tier.label_for(i, values[i]), False)
               for i in range(len(t.categories))]
    total = t.summary_row("AC")
    entries.append((L["row_ac"], total.variance_rel,
                    f"{total.variance_rel:+.0f}", True))

    for y, value, text, bold in entries:
        colour = S.variance_colour(value, tier.higher_is_better)
        x = zero + value * scale
        left, right = sorted((x, zero))
        c.rect(left, y - L["pin_stem"] / 2, max(right - left, 1.0), L["pin_stem"],
               fill=colour)
        fill, outline = paint_for(c, "AC")
        hx = x + (0 if value >= 0 else -head)
        c.rect(hx, y - head / 2, head, head, fill=fill,
               stroke=outline or fill, stroke_width=0.8)
        weight = "bold" if bold else "normal"
        if value >= 0:
            c.text(hx + head + 5, y + 4, text, size=F["value"], weight=weight)
        else:
            c.text(hx - 5, y + 4, text, size=F["value"], anchor="end", weight=weight)

    c.text(L["rel_header_x"], L["header_y"], tier.label, size=F["header"],
           anchor="middle")


def _c06_labels(c: Canvas, t: D.Template, L: dict, rows: _C06Rows) -> None:
    F = L["font"]
    for i, name in enumerate(t.categories):
        c.text(L["label_right"], rows.centre(i) + 4, name, size=F["category"],
               anchor="end")


def _c06_annotations(c: Canvas, t: D.Template, L: dict, rows: _C06Rows) -> None:
    """The circled total, and the arrow marking where growth turns into decline."""
    row = t.summary_row("ΔPY")
    lo, hi = (t.summary_row(name).total for name in row.span)
    text = f"{row.variance_abs:+.0f}"
    cx = (_c06_x(L, lo) + _c06_x(L, hi)) / 2
    half = _text_width(text, L["font"]["value"], False) / 2 + 9
    c.text(cx, L["wf_total_y"] + 20, text, size=L["font"]["value"], anchor="middle")
    highlight_oval(c, cx, L["wf_total_y"] + 15.5, half, 11)

    for note in t.annotations:
        if note.kind != "arrow":
            continue
        _, i = note.target
        spans = D.waterfall_spans(t, t.tier("wf").series_for("AC").values,
                                  start=t.summary_row("PY").total)
        x = _c06_x(L, spans[i][1])
        y = rows.centre(i)
        blue = S.ANNOTATION["highlight"]
        c.line(x, y - 34, x, y - 8, stroke=blue, stroke_width=3)
        c.add(f'<polygon points="{x:.2f},{y - 2:.2f} {x - 5:.2f},{y - 12:.2f} '
              f'{x + 5:.2f},{y - 12:.2f}" fill="{blue}"/>')


# --------------------------------------------------------------------------- #
# C05X layout - measured from template-refs/C05_05X.png
# --------------------------------------------------------------------------- #
#
# The vertical counterpart of C06F: a bridge, but running left to right across
# twelve months rather than down a list of states, and with its last four steps
# forecast. Everything below the pin tier shares one baseline and one scale -
# the two opening columns, the twelve monthly columns, the waterfall and the
# closing column are all 1.5975 px per kEUR, which is what lets a month's step
# be compared with the month's own column directly beneath it.
#
# The category midpoint carries the plan bar and the pin; the actual sits 7px to
# its right, so the notation reads plan-behind-actual the way C03A reads
# prior-year-behind-actual.

C05X_LAYOUT = dict(
    page=(1280, 720),
    margin_left=20,
    margin_right=20,
    message_width=545,
    title_rule_y=76.5,

    baseline=635.5,
    unit_scale=1.5975,            # px per kEUR, shared by every column and step

    cat_x0=300.0,                 # centre of January
    cat_pitch=56.0,
    bar_width=39.0,
    ac_offset=7.0,                # the actual sits right of the plan behind it

    # The three columns that are not months.
    summary_x={"2024 AC": 115.5, "2025 PL": 203.5, "2025 AC+FC": 1027.5},
    summary_width=50.0,
    # The two closing variance bars start at the same x and are drawn to
    # different widths - measured, and it is not decoration: the wide one sits
    # along the prior-year guide it is measured from, and the narrow one drops
    # down the plan guide, so each reads against the line it belongs to.
    summary_var_x=1098.0,
    summary_var_width={"vs PY": 64.0, "vs PL": 8.0},

    # Guides at the prior-year and plan levels, so the bridge can be read
    # against both of the columns it is drawn between.
    guide_x=(91.0, 1160.0),

    divider_x=726.5,              # where measured becomes expected
    divider_top=92.0,

    pin_axis_y=236.25,
    pin_scale=1.572,              # px per percentage point - its own unit
    pin_stem=5,
    pin_head=10,
    pin_label_gap=7,

    step_height_min=3.0,
    label_gap=5.0,

    font=dict(entity=13, measure=13.5, period=13, message=13.5, badge=34,
              value=13, category=13, header=13, total=13.5, footer=9),
)


class _C05Cats:
    """The shared category axis. The months are evenly pitched; the three
    summary columns sit outside the pitch at measured positions, because they
    are aggregates rather than another month."""

    def __init__(self, L: dict, count: int) -> None:
        self.L, self.count = L, count

    def centre(self, i: int) -> float:
        return self.L["cat_x0"] + self.L["cat_pitch"] * i

    def actual(self, i: int) -> float:
        return self.centre(i) + self.L["ac_offset"]


def render_c05x(template: D.Template = D.C05X, layout: dict | None = None) -> str:
    """Render C05X - net sales, plan bridged to actual plus forecast."""
    L = {**C05X_LAYOUT, **(layout or {})}
    t = template
    width, height = L["page"]
    c = Canvas(width, height, prefix=f"{t.id}{t.variant}".lower())
    cats = _C05Cats(L, len(t.categories))

    _title_block(c, t, L)
    _c05_guides(c, t, L)
    _c05_summary_columns(c, t, L)
    _c05_month_columns(c, t, L, cats)
    _c05_waterfall(c, t, L, cats)
    _c05_summary_variances(c, t, L)
    _c05_pins(c, t, L, cats)
    _c05_baseline(c, t, L, cats)
    _c05_divider(c, t, L)
    _c05_annotations(c, t, L, cats)
    _footer(c, L)
    return c.to_svg()


def _c05_y(L: dict, value: float) -> float:
    return L["baseline"] - value * L["unit_scale"]


def _c05_guides(c: Canvas, t: D.Template, L: dict) -> None:
    """A rule at the prior-year and plan levels, run right across the chart.

    Without them the bridge floats: a step is a difference from the plan, and
    the plan is a column three hundred pixels to the left. The rule carries it.
    """
    x0, x1 = L["guide_x"]
    for label in ("2024 AC", "2025 PL"):
        y = _c05_y(L, t.summary_row(label).total)
        c.line(x0, y, x1, y, stroke=S.PAGE["rule"], stroke_width=1)


def _c05_summary_columns(c: Canvas, t: D.Template, L: dict) -> None:
    """The opening columns and the closing stack.

    The closing column is the only place in the template where measured and
    expected are added together, and it says so by stacking them rather than
    totalling them: solid below, hatched above, each labelled with its own part.
    """
    F, w = L["font"], L["summary_width"]
    for label, x in L["summary_x"].items():
        row = t.summary_row(label)
        base = L["baseline"]
        for scenario, value in row.stack:
            top = base - value * L["unit_scale"]
            fill, outline = paint_for(c, scenario)
            c.rect(x - w / 2, top, w, base - top, fill=fill,
                   stroke=outline or "none",
                   stroke_width=1 if outline else 0)
            # A stacked part is labelled inside itself, because its number is
            # not the column's total and putting it on top would say it was.
            if len(row.stack) > 1:
                c.text(x + w / 2 - 6, (top + base) / 2 + 4, f"{value:,.0f}",
                       size=F["value"], anchor="end",
                       fill="#FFFFFF" if scenario == "AC" else S.TEXT["primary"])
                c.text(x + w / 2 + 8, (top + base) / 2 + 4, scenario,
                       size=F["value"])
            base = top
        c.text(x, _c05_y(L, row.total) - 8, f"{row.total:,.0f}",
               size=F["total"], anchor="middle")


def _c05_month_columns(c: Canvas, t: D.Template, L: dict, cats: _C05Cats) -> None:
    """Plan behind, actual or forecast in front, and the fill says which.

    September onwards is hatched. That is the whole point of the template: the
    reader can see, without a legend, exactly where the year stops being
    measured and starts being expected.
    """
    tier = t.tier("measure")
    pl = tier.series_for("PL").values
    F, w, base = L["font"], L["bar_width"], L["baseline"]

    for i, scenario in enumerate(t.category_scenarios):
        value = _value(tier, scenario, i)
        if value is None:
            continue
        scenario_bar(c, "PL", cats.centre(i), w, _c05_y(L, pl[i]), base)
        scenario_bar(c, scenario, cats.actual(i), w, _c05_y(L, value), base)
        c.text(cats.actual(i), _c05_y(L, value) - 7, f"{value:,.0f}",
               size=F["value"], anchor="middle")


def _c05_waterfall(c: Canvas, t: D.Template, L: dict, cats: _C05Cats) -> None:
    """The bridge, one step per month, starting on the plan column.

    Each step is coloured by impact and filled by scenario: a forecast step is
    hatched in its own variance colour, so an expected improvement cannot be
    mistaken for a measured one. That combination - impact deciding the colour,
    scenario deciding the fill - is the densest piece of notation in the
    library, and it is why C05 is worth building.
    """
    tier = t.tier("wf")
    values = tier.series_for("AC").values
    spans = D.waterfall_spans(t, values, start=t.summary_row("2025 PL").total)
    F, w = L["font"], L["bar_width"]

    for i, (lo, hi) in enumerate(spans):
        scenario = t.category_scenarios[i]
        colour = S.variance_colour(values[i], tier.higher_is_better)
        y0, y1 = sorted((_c05_y(L, lo), _c05_y(L, hi)))
        x = cats.actual(i) - w / 2
        height = max(y1 - y0, L["step_height_min"])
        if scenario == "FC":
            fill = c.hatch(S.Hatch(colour, 45, "ascending", 8, 3),
                           S.scenario_fill("FC").fill)
            c.rect(x, y0, w, height, fill=fill, stroke=colour, stroke_width=1)
        else:
            c.rect(x, y0, w, height, fill=colour)

        text = tier.label_for(i, values[i])
        if values[i] >= 0:
            c.text(cats.actual(i), y0 - 6, text, size=F["value"], anchor="middle")
        else:
            c.text(cats.actual(i), y1 + 13, text, size=F["value"], anchor="middle")


def _c05_summary_variances(c: Canvas, t: D.Template, L: dict) -> None:
    """The closing column against each of the two it is drawn beside."""
    F = L["font"]
    x = L["summary_var_x"]
    for row in t.summary_rows:
        if not row.is_variance_only:
            continue
        w = L["summary_var_width"][row.label]
        lo, hi = (t.summary_row(name).total for name in row.span)
        y0, y1 = sorted((_c05_y(L, lo), _c05_y(L, hi)))
        colour = S.variance_colour(row.variance_abs, True)
        c.rect(x, y0, w, max(y1 - y0, L["step_height_min"]), fill=colour)
        c.text(x + w + 6, (y0 + y1) / 2 + 4, f"{row.variance_abs:+.0f}",
               size=F["value"])


def _c05_pins(c: Canvas, t: D.Template, L: dict, cats: _C05Cats) -> None:
    """Relative variances as pins on a plan axis, which is a double rule.

    The head carries the scenario, so a forecast pin's head is hatched - the
    same distinction the columns make, made again where a reader is looking at
    percentages rather than amounts.
    """
    tier = t.tier("var_rel")
    values = tier.series_for("AC").values
    zero, scale, F = L["pin_axis_y"], L["pin_scale"], L["font"]
    x0, x1 = L["guide_x"]

    reference_axis(c, tier.reference, x0 + 182, x1 - 100, zero)

    entries = [(cats.centre(i), values[i], t.category_scenarios[i],
                tier.label_for(i, values[i])) for i in range(len(values))]
    closing = t.summary_row("2025 AC+FC")
    entries.append((L["summary_x"]["2025 AC+FC"], closing.variance_rel, "FC",
                    f"{closing.variance_rel:+.1f}"))

    for cx, value, scenario, text in entries:
        colour = S.variance_colour(value, tier.higher_is_better)
        y = zero - value * scale
        variance_pin(c, scenario, cx, L["pin_stem"], y, zero, colour)
        above = value >= 0
        c.text(cx, y - L["pin_label_gap"] - 4 if above else y + L["pin_label_gap"] + 12,
               text, size=F["value"], anchor="middle")

    c.text(x0 + 176, zero + 4, tier.label, size=F["header"], anchor="end")


def _c05_baseline(c: Canvas, t: D.Template, L: dict, cats: _C05Cats) -> None:
    """The category axis, and the two-line labels under it."""
    F, base = L["font"], L["baseline"]
    c.line(80, base, 1063, base, stroke=S.TEXT["primary"], stroke_width=1.6)

    for i, month in enumerate(t.categories):
        c.text(cats.centre(i) + L["ac_offset"] / 2, base + 17, month,
               size=F["category"], anchor="middle")
    # Each label is two lines: what it is, then which scenario it is in. The
    # second line is what stops "2025" appearing three times and meaning three
    # different things.
    for label, x in L["summary_x"].items():
        head, _, tail = label.partition(" ")
        c.text(x, base + 17, head, size=F["category"], anchor="middle")
        c.text(x, base + 33, tail, size=F["category"], anchor="middle")
    for first, last, tail in ((0, 7, "AC"), (8, 11, "FC")):
        c.text((cats.centre(first) + cats.centre(last)) / 2 + L["ac_offset"] / 2,
               base + 33, tail, size=F["category"], anchor="middle")


def _c05_divider(c: Canvas, t: D.Template, L: dict) -> None:
    """The line where measured becomes expected."""
    c.line(L["divider_x"], L["divider_top"], L["divider_x"], L["baseline"] + 14,
           stroke=S.TEXT["primary"], stroke_width=1)


def _c05_annotations(c: Canvas, t: D.Template, L: dict, cats: _C05Cats) -> None:
    """Ovals round the two numbers the message names."""
    closing = t.summary_row("2025 AC+FC")
    cx = L["summary_x"]["2025 AC+FC"]
    y = L["pin_axis_y"] - closing.variance_rel * L["pin_scale"]
    text = f"{closing.variance_rel:+.1f}"
    half = _text_width(text, L["font"]["value"], False) / 2 + 8
    highlight_oval(c, cx, y - L["pin_label_gap"] - 8, half, 11)

    row = t.summary_row("vs PL")
    lo, hi = (t.summary_row(name).total for name in row.span)
    y0, y1 = sorted((_c05_y(L, lo), _c05_y(L, hi)))
    text = f"{row.variance_abs:+.0f}"
    half = _text_width(text, L["font"]["value"], False) / 2 + 8
    x = L["summary_var_x"] + L["summary_var_width"]["vs PL"] + 6
    highlight_oval(c, x + half - 8, (y0 + y1) / 2, half, 11)


# --------------------------------------------------------------------------- #
# The table engine
# --------------------------------------------------------------------------- #
#
# Nothing above this line helps with a table. A chart maps values to lengths; a
# table maps them to a grid of text, and what has to be got right is alignment,
# hierarchy and the notation carried by the column headers. So this is a second
# small engine rather than a new kind of chart - but it reuses the data layer
# whole, because a column of a table turns out to be a Tier: a keyed set of
# values, optionally measured against a named scenario.
#
# Three ideas and the rest follows:
#
#   TableGrid     rows at a uniform pitch, with a gap after each group subtotal
#   TableColumn   one printed column - where it sits, and which series it prints
#   scenario_rule the header rule that says what a column holds, in the same
#                 notation a bar would use: solid dark AC, solid light PY, a
#                 double rule for the outlined PL
#
# A panel column - one whose cells are drawn rather than typed - is a
# TableColumn like any other; T02A and T04A hang the existing variance
# primitives inside the cell rectangle this hands them.


@dataclass(frozen=True)
class TableGrid:
    """Where each row of a table sits.

    Rows are a uniform pitch with one exception, which is structural rather
    than decorative: a group subtotal ends a group, so the row after it starts
    a new one and is pushed down by ``group_gap``. Deriving that from
    ``Row.kind`` rather than from a list of indices means the spacing cannot
    disagree with the hierarchy the subtotals actually describe.
    """

    y0: float
    row_height: float
    group_gap: float
    rows: Sequence[D.Row]
    # Where a statement breaks its sections, when that is not simply "after
    # every subtotal". T03A rules off after sales revenue and after each result
    # line but not after operating expenses, and leaves a half gap under the
    # margin - editorial choices in the source, so they are declared rather
    # than derived from the hierarchy, which does not know about them.
    gaps: dict | None = None

    @property
    def tops(self) -> list[float]:
        out, y = [], self.y0
        for i, row in enumerate(self.rows):
            if self.gaps is not None:
                y += self.gaps.get(i - 1, 0.0)
            elif i and self.rows[i - 1].kind == "subtotal":
                y += self.group_gap
            out.append(y)
            y += self.row_height
        return out

    def top(self, index: int) -> float:
        return self.tops[index]

    def bottom(self, index: int) -> float:
        return self.tops[index] + self.row_height

    def centre(self, index: int) -> float:
        return self.tops[index] + self.row_height / 2

    @property
    def height(self) -> float:
        return self.tops[-1] + self.row_height - self.y0


@dataclass(frozen=True)
class TableColumn:
    """One printed column of a table.

    ``series_index`` picks which of the tier's series this column prints: a
    measure tier holds PY, PL and AC and spends three columns doing it, while a
    variance tier holds one series and spends one.
    """

    x0: float
    width: float
    tier: D.Tier
    series_index: int = 0

    @property
    def x1(self) -> float:
        return self.x0 + self.width

    @property
    def series(self) -> D.Series:
        return self.tier.series[self.series_index]

    @property
    def scenario(self) -> str:
        return self.series.scenario

    @property
    def caption(self) -> str:
        """What the column header says, or nothing if a neighbour says it.

        A measure column falls back to its scenario name - that *is* its
        header. A variance column does not: an empty label means this is the
        second half of a pair whose first half carries the heading, and falling
        back would head a ΔPY% column "AC".
        """
        if self.tier.kind == "measure":
            return self.tier.label or self.series.label or self.series.scenario
        return self.tier.label

    def value(self, row: int) -> float | None:
        return self.series.values[row]


def table_columns(t: D.Template, L: dict) -> tuple[list[TableColumn], dict, float]:
    """Turn the data layer's column order into x positions.

    The *order* is not decided here - ``D.table_column_plan`` decides it, and
    the worksheet builder reads the same plan. This only turns each entry's gap
    class into pixels, so the two engines cannot lay a table out differently.
    """
    gutter, group_gutter = L["col_gutter"], L["group_gutter"]
    widths = L.get("col_width_for", {})
    plan = D.table_column_plan(t, L["label_after_block"])

    columns: list[TableColumn] = []
    block_extent: dict[str, list[float]] = {}
    label_x = None
    x = L["blocks_x0"]

    per_column = L.get("col_gutter_for", {})
    for entry in plan:
        if entry.gap == "same":
            key = entry.tier.key if entry.kind == "value" else None
            x += per_column.get(key, gutter)
        elif entry.gap == "group":
            x += group_gutter
        if entry.kind == "label":
            label_x = x
            x += L["label_width"]
            continue
        # A panel column is wider than a printed one, and the two panels of a
        # block are not the same width as each other: the width is what absorbs
        # a larger range once the scale is fixed.
        width = widths.get(entry.tier.key, L["col_width"])
        columns.append(TableColumn(x, width, entry.tier, entry.series_index))
        extent = block_extent.setdefault(entry.tier.block, [x, x])
        extent[1] = x + width
        x += width

    return columns, {k: tuple(v) for k, v in block_extent.items()}, label_x


def scenario_rule(canvas: Canvas, scenario: str, x0: float, x1: float,
                  y: float, weight: float) -> None:
    """The header rule that says which scenario a column holds.

    The same statement a bar makes with its fill, made by a rule instead: solid
    dark for an actual, solid light for a prior year, two thin rules for the
    outlined plan, hatched for a forecast. A reader who has learned the fills
    has already learned these, which is the whole point of the notation.

    ``reference_axis`` says the same thing for a variance tier's axis, but only
    for the scenario being compared *against* - it rejects AC by design. A table
    column header has to be able to say "these are actuals", so this is the
    general form.
    """
    spec = S.scenario_fill(scenario)
    if spec.hatch:
        canvas.rect(x0, y, x1 - x0, weight,
                    fill=canvas.hatch(spec.hatch, spec.fill))
    elif spec.is_fictitious:
        # Outlined, seen edge-on: the top and bottom of the outline, and the
        # empty middle that says nobody measured this.
        thin = weight * 0.36
        canvas.rect(x0, y, x1 - x0, thin, fill=spec.outline)
        canvas.rect(x0, y + weight - thin, x1 - x0, thin, fill=spec.outline)
    else:
        canvas.rect(x0, y, x1 - x0, weight, fill=spec.fill)


def _table_header(c: Canvas, t: D.Template, L: dict, columns: list[TableColumn],
                  block_extent: dict, label_x: float) -> None:
    """Block headers, column captions, and the rule under each."""
    F = L["font"]

    for block, (x0, x1) in block_extent.items():
        if not block:
            continue        # a single-block table has nothing to head
        c.text((x0 + x1) / 2, L["block_header_baseline"], block,
               size=F["caption"], anchor="middle")
        c.rect(x0, L["block_rule_y"], x1 - x0, L["block_rule_weight"],
               fill=S.TEXT["primary"])

    for column in columns:
        # A printed column's caption sits over its digits, so it is right
        # aligned like them. A panel's sits over the panel, centred - not over
        # the zero rule, which on an absolute panel is not the middle.
        if not column.caption:
            continue        # the second half of a pair; the first carries it
        if column.tier.key in t.panel_tiers:
            c.text((column.x0 + column.x1) / 2, L["caption_baseline"],
                   column.caption, size=F["caption"], anchor="middle")
        else:
            # A pair - dPY over dPY and dPY% - is headed once, at the far edge
            # of the pair rather than of its own column. Derived from the
            # neighbours rather than declared: an unlabelled column measured
            # against the same scenario is the second half of a pair.
            right = column.x1
            for other in columns[columns.index(column) + 1:]:
                if other.caption or other.tier.reference != column.tier.reference:
                    break
                right = other.x1
            c.text(right - L["cell_pad"], L["caption_baseline"],
                   column.caption, size=F["caption"], anchor="end")
        if column.tier.kind == "measure":
            scenario_rule(c, column.scenario, column.x0, column.x1,
                          L["scenario_rule_y"], L["scenario_rule_weight"])
        else:
            c.rect(column.x0, L["caption_rule_y"], column.width,
                   L["caption_rule_weight"], fill=S.TEXT["primary"])

    # The label column carries the same rule as a variance column, so the header
    # band reads as one line across the page rather than stopping at each block.
    c.rect(label_x, L["caption_rule_y"], L["label_width"],
           L["caption_rule_weight"], fill=S.TEXT["primary"])


def _table_body(c: Canvas, t: D.Template, L: dict, grid: TableGrid,
                columns: list[TableColumn], label_x: float) -> None:
    """Every cell, plus the row labels."""
    F = L["font"]
    ratio_formats = L.get("ratio_formats", {})
    for i, row in enumerate(t.rows):
        bold = "bold" if row.kind == "subtotal" else "normal"
        italic = "italic" if row.kind == "ratio" else "normal"
        baseline = grid.top(i) + L["baseline_offset"]

        indent = row.indent * L.get("indent_step", 0.0)
        prefix = row.prefix
        if prefix:
            c.text(label_x + L["label_pad"], baseline, prefix,
                   size=F["cell"], weight=bold)
        c.text(label_x + L["label_pad"] + L.get("label_indent", 0.0) + indent,
               baseline, row.label, size=F["cell"], weight=bold, style=italic)

        for column in columns:
            if column.tier.key in t.panel_tiers:
                _panel_cell(c, t, L, grid, column, i)
                continue
            value = column.value(i)
            if value is None:
                continue
            red = (column.tier.red_below is not None
                   and value < column.tier.red_below)
            # A ratio row is not in the column's unit: its figures are
            # percentages and its absolute variance is in percentage *points*,
            # which is what the "%p" says. Same numbers, different denomination.
            if row.kind == "ratio" and column.tier.kind in ratio_formats:
                text = ratio_formats[column.tier.kind].format(value)
            else:
                text = column.tier.label_for(i, value)
            c.text(column.x1 - L["cell_pad"], baseline, text, size=F["cell"],
                   weight=bold, anchor="end", style=italic,
                   fill=S.VARIANCE["bad"] if red else S.TEXT["primary"])


def _table_rules(c: Canvas, t: D.Template, L: dict, grid: TableGrid,
                 columns: list[TableColumn], label_x: float) -> None:
    """The rule under every row, heavier where a subtotal closes a group.

    Drawn per column rather than as one line across the page: the reference
    breaks its rules at the gutters, which is what keeps the eye inside a
    column when reading down it.
    """
    spans = [(column.x0, column.width) for column in columns]
    spans.append((label_x, L["label_width"]))

    for i, row in enumerate(t.rows):
        # A group subtotal is ruled off above, where the group ends; every other
        # row takes the light separator below it. The World row is a subtotal
        # like the others, so it needs no special case.
        if row.kind == "subtotal":
            y, weight, fill = (grid.top(i) - L["subtotal_rule_lift"],
                               L["subtotal_rule_weight"], S.TEXT["primary"])
        elif i + 1 < len(t.rows) and t.rows[i + 1].kind == "subtotal":
            continue        # the subtotal's own rule already closes this row
        else:
            y, weight, fill = (grid.bottom(i) - L["row_rule_lift"],
                               L["row_rule_weight"], L["row_rule_colour"])
        for x0, width in spans:
            c.rect(x0, y, width, weight, fill=fill)


def _table_footnote(c: Canvas, L: dict, columns: list[TableColumn]) -> None:
    """The threshold under each variance column, and the phrase that names it.

    IBCS requires a threshold to be stated wherever it is applied. Printing it
    from ``Tier.red_below`` - the same number the cells are coloured by - is
    what makes the statement true rather than a caption that used to be true.
    """
    F, grey = L["font"], S.PAGE["footnote"]
    marked = [column for column in columns if column.tier.red_below is not None]
    if not marked:
        return

    c.text(marked[0].x0 - L["group_gutter"], L["footnote_baseline"],
           "Δ in red if:", size=F["footnote"], fill=grey, anchor="end",
           style="italic")
    for column in marked:
        threshold = column.tier.red_below
        text = (f"<{threshold:,.0f}" if column.tier.kind == "variance_abs"
                else f"<{threshold:,.0f}%")
        c.text(column.x1 - L["cell_pad"], L["footnote_baseline"], text,
               size=F["footnote"], fill=grey, anchor="end", style="italic")


def panel_geometry(column: "TableColumn", L: dict) -> tuple[float, float]:
    """A panel column's zero position and its scale, in px per unit.

    Both are measured off the reference rather than fitted to the data, and the
    scale is shared by every panel measuring the same unit - which is what makes
    the two blocks comparable. T02A's month and cumulative ΔPL panels carry the
    same 0.447 px per kEUR and differ in *width* instead, because the cumulative
    figures are the larger ones. Fitting each panel to its own range would have
    drawn a month variance of 88 the same length as a cumulative one of 211.
    """
    geometry = L["panels"][column.tier.key]
    return column.x0 + geometry.zero_px, geometry.scale


def _panel_cell(c: Canvas, t: D.Template, L: dict, grid: TableGrid,
                column: "TableColumn", index: int) -> None:
    """One drawn variance: the element, then its label outside the far end.

    An absolute variance is a bar and a relative one is a pin - the same rule
    as every chart in the library, and the reason a reader can move between a
    T02 table and a C04 chart without relearning anything.
    """
    value = column.value(index)
    if value is None:
        return
    zero, scale = panel_geometry(column, L)
    cy = grid.centre(index)
    x_value = zero + value * scale
    colour = S.variance_colour(value, t.higher_is_better_at(index, column.tier))

    if column.tier.kind == "variance_rel":
        variance_pin_h(c, "AC", cy, L["pin_stem"], x_value, zero, colour)
        reach = max(L["pin_stem"] * S.PIN_HEAD_RATIO / 2, 0)
    else:
        left, right = sorted((zero, x_value))
        c.rect(left, cy - L["bar_height"] / 2, right - left, L["bar_height"],
               fill=colour)
        reach = 0

    text = column.tier.label_for(index, value)
    bold = "bold" if t.rows[index].kind == "subtotal" else "normal"
    gap = L["panel_label_gap"] + reach
    if value < 0:
        c.text(x_value - gap, grid.top(index) + L["baseline_offset"], text,
               size=L["font"]["cell"], anchor="end", weight=bold)
    else:
        c.text(x_value + gap, grid.top(index) + L["baseline_offset"], text,
               size=L["font"]["cell"], anchor="start", weight=bold)


def _panel_axes(c: Canvas, t: D.Template, L: dict, grid: TableGrid,
                columns: list["TableColumn"]) -> None:
    """The zero rule of each panel, in the notation of its reference scenario.

    This is the vertical form of the same statement a variance tier's axis
    makes: a double rule for a plan comparison, because plan is drawn outlined
    and an outlined bar seen edge-on is two lines. It is why these panels need
    no legend to say what they are variances from.
    """
    y0 = L["caption_rule_y"] + L["caption_rule_weight"]
    y1 = grid.bottom(len(t.rows) - 1)
    for column in columns:
        if column.tier.key not in t.panel_tiers:
            continue
        zero, _ = panel_geometry(column, L)
        reference_axis_v(c, column.tier.reference, y0, y1, zero)


def render_table(t: D.Template, L: dict) -> str:
    """Render any of the four table templates.

    One function for all four: T01B and T03A print every cell, while T02A and
    T04A replace two of their columns with drawn panels. That is a difference in
    what a cell contains, not in how the table is built, so it is handled by the
    panel columns rather than by a second renderer.
    """
    c = Canvas(*L["page"], prefix=f"{t.id}{t.variant}".lower())
    columns, block_extent, label_x = table_columns(t, L)
    grid = TableGrid(L["row0_top"], L["row_height"], L["group_gap"], t.rows,
                     L.get("row_gaps"))

    _title_block(c, t, L)
    _table_header(c, t, L, columns, block_extent, label_x)
    _table_rules(c, t, L, grid, columns, label_x)
    _panel_axes(c, t, L, grid, columns)
    _table_body(c, t, L, grid, columns, label_x)
    _table_footnote(c, L, columns)
    _footer(c, L)
    return c.to_svg()


# --------------------------------------------------------------------------- #
# T01B layout - measured from template-refs/T01_T01B.png
# --------------------------------------------------------------------------- #
#
# The grid is exactly regular once measured, which is worth saying because it
# looks irregular: columns are 64 wide on an 8px gutter inside a group and a
# 16px gutter between groups, and the row labels are a 112-wide column sitting
# between the two blocks rather than at either edge.
#
# Rows are 24 apart from y=144, with a 12px gap after each group subtotal. The
# type is one size throughout - 15.5px, solved from "Switzerland" at 80px wide
# and confirmed on "November" at 72 - and the only weight change is bold on the
# subtotal rows.

T01B_LAYOUT = dict(
    page=(1280, 720),
    margin_left=20,
    margin_right=20,
    title_rule_y=76.5,
    message_width=520,

    blocks_x0=48,
    col_width=64,
    col_gutter=8,           # between columns measured against the same reference
    group_gutter=16,        # between one reference and the next
    label_width=112,
    label_after_block=1,    # the row labels sit between the two blocks
    cell_pad=4.5,
    label_pad=3.5,

    row0_top=144.0,
    row_height=24.0,
    # 11, not 12: solved from the 131px between consecutive subtotal rules,
    # which is four rows plus one gap.
    group_gap=11.0,
    baseline_offset=19.5,

    block_header_baseline=107.5,
    block_rule_y=111.0,
    block_rule_weight=1.6,
    caption_baseline=131.5,
    caption_rule_y=135.4,
    caption_rule_weight=1.6,
    # The three scenario rules are the notation: 4.5px of solid grey for PY,
    # solid dark for AC, and two thin rules for the outlined PL.
    scenario_rule_y=136.2,
    scenario_rule_weight=4.6,

    row_rule_lift=0.0,
    row_rule_weight=1.0,
    row_rule_colour="#BDBDBD",
    subtotal_rule_lift=0.9,      # the rule straddles the row top, not above it
    subtotal_rule_weight=1.8,

    footnote_baseline=671.5,
    font=dict(entity=13, measure=13.5, period=13, message=13.5, badge=34,
              caption=15.5, cell=15.5, footnote=12.5, footer=9),
)


def render_t01b(template: D.Template = D.T01B, layout: dict | None = None) -> str:
    """Render T01B - Pharmaceutical Inc. profit after tax, month and year to date."""
    return render_table(template, {**T01B_LAYOUT, **(layout or {})})


# --------------------------------------------------------------------------- #
# T02A layout - measured from template-refs/T02_T02A.png
# --------------------------------------------------------------------------- #
#
# The same table as T01B with its variance columns drawn, so the header band and
# the row pitch are shared and only three things differ: the columns are packed
# tighter (5px gutters inside a group, 9 between), the groups are 16 apart
# rather than 11, and four of the columns are panels.
#
# The panel scales are the interesting measurement. Both ΔPL panels carry the
# same 0.4472 px per kEUR and both ΔPL% panels the same 3.60 px per point - the
# month and the cumulative block are directly comparable - and the cumulative
# ΔPL panel is twice as wide because its numbers are twice as big. Scale fixed,
# width variable, which is the right way round: a panel fitted to its own range
# would draw a variance of 88 and one of 211 the same length.

T02A_LAYOUT = dict(
    page=(1280, 720),
    margin_left=20,
    margin_right=20,
    title_rule_y=76.5,
    message_width=520,

    blocks_x0=32,
    col_width=64,
    col_gutter=5,
    group_gutter=9,
    label_width=128,
    label_after_block=1,
    cell_pad=4.5,
    label_pad=3.5,
    # Widths, zero positions and scales all come from the shared measurement,
    # so the worksheet cannot put a zero rule somewhere the page does not.
    panels=T02A_PANELS,
    col_width_for={k: g.width_px for k, g in T02A_PANELS.items()},
    bar_height=11.0,
    pin_stem=4.0,
    panel_label_gap=5.0,

    row0_top=144.0,
    row_height=24.0,
    group_gap=16.0,
    baseline_offset=18.0,

    block_header_baseline=107.5,
    block_rule_y=111.0,
    block_rule_weight=1.6,
    caption_baseline=131.5,
    caption_rule_y=135.4,
    caption_rule_weight=1.6,
    scenario_rule_y=136.2,
    scenario_rule_weight=4.6,

    row_rule_lift=0.0,
    row_rule_weight=1.0,
    row_rule_colour="#BDBDBD",
    subtotal_rule_lift=0.9,
    subtotal_rule_weight=1.8,

    footnote_baseline=690.0,
    font=dict(entity=13, measure=13.5, period=13, message=13.5, badge=34,
              caption=15.5, cell=15.5, footnote=12.5, footer=9),
)


def render_t02a(template: D.Template = D.T02A, layout: dict | None = None) -> str:
    """Render T02A - the same report as T01B, with its variances drawn."""
    return render_table(template, {**T02A_LAYOUT, **(layout or {})})



# --------------------------------------------------------------------------- #
# T03A layout - measured from template-refs/T03_T03A.png
# --------------------------------------------------------------------------- #
#
# The same grid as T01B with the labels leading rather than sitting between two
# blocks, and one measured irregularity worth naming: the scenario columns are 72
# wide on a 3px gutter, but each variance *pair* is two 64-wide columns butted
# together with no gutter at all, headed once at the far edge of the pair. The
# consistent 5px of padding on both halves is what settles it - a 62/4 split
# would put the padding at 3 and 5.
#
# The row gaps are editorial rather than structural. The statement breaks after
# sales revenue and after each result line, but not after operating expenses,
# and leaves a half gap under the margin - so they are declared, not derived.

T03A_LAYOUT = dict(
    page=(1280, 720),
    margin_left=20,
    margin_right=20,
    title_rule_y=76.5,
    message_width=520,

    blocks_x0=48,
    col_width=72,
    col_gutter=3,
    group_gutter=13,
    label_width=192,
    label_after_block=0,
    col_width_for={"dpy": 64, "dpyp": 64, "dpl": 64, "dplp": 64},
    col_gutter_for={"dpyp": 0.0, "dplp": 0.0},
    cell_pad=4.5,
    label_pad=4.0,
    label_indent=16.0,      # the +/-/= prefix sits left of every label
    indent_step=9.0,        # a component line, indented under the line it feeds

    row0_top=117.0,
    row_height=24.0,
    group_gap=11.0,
    row_gaps={4: 11.0, 13: 6.0, 16: 11.0, 18: 11.0},
    baseline_offset=18.5,

    block_header_baseline=0.0,      # unused: one block, so no header band
    block_rule_y=0.0,
    block_rule_weight=0.0,
    caption_baseline=103.5,
    caption_rule_y=108.4,
    caption_rule_weight=1.6,
    scenario_rule_y=109.0,
    scenario_rule_weight=4.6,

    row_rule_lift=0.0,
    row_rule_weight=1.0,
    row_rule_colour="#BDBDBD",
    subtotal_rule_lift=0.9,
    subtotal_rule_weight=1.8,

    # A ratio row is not in the column's unit.
    ratio_formats={"measure": "{:,.1f}%",
                   "variance_abs": "{:+,.1f}%p",
                   "variance_rel": "{:+,.0f}%"},

    footnote_baseline=690.0,
    font=dict(entity=13, measure=13.5, period=13, message=13.5, badge=34,
              caption=15.5, cell=15.5, footnote=12.5, footer=9),
)


def render_t03a(template: D.Template = D.T03A, layout: dict | None = None) -> str:
    """Render T03A - a full P&L against both prior year and plan."""
    return render_table(template, {**T03A_LAYOUT, **(layout or {})})



# --------------------------------------------------------------------------- #
# T04A layout - measured from template-refs/T04_T04A.png
# --------------------------------------------------------------------------- #
#
# T03A's statement with T02A's panels, so both halves are already built and this
# is measurement rather than mechanism. The panels are wider than T02A's and far
# more finely scaled - 5.71 px per kEUR against 0.45 - because a statement's
# variances are single digits where a country table's run to hundreds.

T04A_LAYOUT = dict(
    page=(1280, 720),
    margin_left=20,
    margin_right=20,
    title_rule_y=76.5,
    message_width=520,

    blocks_x0=69,
    col_width=80,
    col_gutter=4,
    group_gutter=5,
    label_width=224,
    label_after_block=0,
    panels=T04A_PANELS,
    col_width_for={k: g.width_px for k, g in T04A_PANELS.items()},
    cell_pad=4.5,
    label_pad=4.0,
    label_indent=16.0,
    indent_step=9.0,
    bar_height=11.0,
    pin_stem=4.0,
    panel_label_gap=5.0,

    row0_top=137.0,
    row_height=24.0,
    group_gap=7.0,
    row_gaps={4: 7.0, 13: 4.0, 16: 6.0, 18: 7.0},
    baseline_offset=18.5,

    block_header_baseline=0.0,
    block_rule_y=0.0,
    block_rule_weight=0.0,
    caption_baseline=126.5,
    caption_rule_y=132.4,
    caption_rule_weight=1.6,
    scenario_rule_y=133.0,
    scenario_rule_weight=4.6,

    row_rule_lift=0.0,
    row_rule_weight=1.0,
    row_rule_colour="#BDBDBD",
    subtotal_rule_lift=0.9,
    subtotal_rule_weight=1.8,

    ratio_formats={"measure": "{:,.1f}%"},

    footnote_baseline=690.0,
    font=dict(entity=13, measure=13.5, period=13, message=13.5, badge=34,
              caption=15.5, cell=15.5, footnote=12.5, footer=9),
)


def render_t04a(template: D.Template = D.T04A, layout: dict | None = None) -> str:
    """Render T04A - a statement with both its variance columns drawn."""
    return render_table(template, {**T04A_LAYOUT, **(layout or {})})

# --------------------------------------------------------------------------- #
# Stacked structure panels
# --------------------------------------------------------------------------- #
#
# A structure chart asks something none of the earlier templates did: the bands
# are categories rather than scenarios, so the fill can no longer say what kind
# of number this is. The scenario moves to the column - four solid actuals and
# an outlined plan - and the bands take a grey ramp instead.
#
# **Divergence, recorded.** The reference does not ramp its greys; it assigns
# them so that no two touching bands are alike, which on the business-area panel
# runs dark, light, mid, lighter, mid-dark from the axis up. That is a choice
# about one chart rather than a rule, and this skill's whole position is that a
# caller asks for "the third structure colour" and never for a hex value - so
# the ramp is used and this note is the difference. The geometry, the totals and
# the notation are reproduced exactly; the shade order is ours.


def _panel_x(L: dict, panel: D.StructurePanel, index: int) -> float:
    """Left edge of one column of one panel."""
    spec = L["panels"][panel.key]
    return spec["x0"] + spec.get("pitch", 0.0) * index


def _segment_fill(canvas: Canvas, panel: D.StructurePanel, position: int,
                  scenario: str) -> tuple[str, str | None]:
    """The fill and outline for one band.

    The band's own colour comes from the structure ramp. What the *scenario*
    does is change how the column is drawn around it: a plan column is outlined
    and its base band goes to the plan fill, because the darkest band is what
    makes a column read as solid, and a solid column reads as measured.
    """
    colour = S.structure_colour(position, len(panel.segments))
    if scenario in ("PL", "BU"):
        spec = S.scenario_fill(scenario)
        if position == 0:
            colour = spec.fill
        return colour, spec.outline
    return colour, None


def _structure_panel(c: Canvas, t: D.Template, L: dict, panel: D.StructurePanel,
                     scale: Scale) -> None:
    """One panel: its columns, the bands in them, and the labels they carry."""
    F, spec = L["font"], L["panels"][panel.key]
    width = L["bar_width"]

    c.text(spec["x0"] + width / 2 + spec.get("header_dx", 0.0),
           L["panel_header_baseline"], panel.label, size=F["header"],
           anchor="middle")

    for index, scenario in enumerate(panel.category_scenarios):
        x = _panel_x(L, panel, index)
        for position, (segment, base, top) in enumerate(panel.stack(index)):
            if top - base <= 0:
                continue
            fill, outline = _segment_fill(c, panel, position, scenario)
            y0, y1 = scale.y(top), scale.y(base)
            c.rect(x, y0, width, y1 - y0, fill=fill, stroke=outline,
                   stroke_width=1 if outline else 0)

            # A label only where the band can hold one. Suppressing it is a
            # required feature of the template, not an omission - a figure
            # crushed into four pixels is worse than no figure.
            if y1 - y0 >= L["label_min_height"]:
                c.text(x + width / 2, (y0 + y1) / 2 + F["value"] * 0.36,
                       f"{segment.at(index):,.1f}", size=F["value"],
                       anchor="middle", fill=S.on_fill(fill))

        total = panel.total(index)
        c.text(x + width / 2, scale.y(total) - L["total_gap"],
               f"{total:,.1f}", size=F["value"], anchor="middle")

    _panel_legend(c, L, panel, scale)
    _panel_categories(c, L, panel)


def _panel_legend(c: Canvas, L: dict, panel: D.StructurePanel,
                  scale: Scale) -> None:
    """Band names beside the bands, which is what saves the reader a legend."""
    F = L["font"]
    x = _panel_x(L, panel, panel.legend_at)
    for segment, base, top in panel.stack(panel.legend_at):
        if top - base <= 0:
            continue
        y = (scale.y(top) + scale.y(base)) / 2 + F["legend"] * 0.36
        if panel.legend_side == "left":
            c.text(x - L["legend_gap"], y, segment.label, size=F["legend"],
                   anchor="end")
        else:
            c.text(x + L["bar_width"] + L["legend_gap"], y, segment.label,
                   size=F["legend"], anchor="start")


def _panel_categories(c: Canvas, L: dict, panel: D.StructurePanel) -> None:
    """The year under each column, and the scenario under the ones that change it.

    The scenario is named where it *starts*, not on every column: writing AC
    five times would say the same thing five times, and the reader only needs
    telling when the answer changes.
    """
    F = L["font"]
    previous = None
    for index, (name, scenario) in enumerate(zip(panel.categories,
                                                 panel.category_scenarios)):
        x = _panel_x(L, panel, index) + L["bar_width"] / 2
        c.text(x, L["category_baseline"], name, size=F["category"],
               anchor="middle")
        if scenario != previous:
            c.text(x, L["scenario_baseline"], scenario, size=F["category"],
                   anchor="middle")
        previous = scenario


def _structure_brackets(c: Canvas, t: D.Template, L: dict,
                        panels: dict, scale: Scale) -> None:
    """The growth brackets the message names.

    Each spans two column tops - or two band tops - and states the difference
    between them. Drawn from the data rather than from the sentence, so a
    retyped figure moves the bracket and the sentence is the thing that has to
    be brought back into line.
    """
    F = L["font"]
    for annotation in t.annotations:
        if annotation.kind != "bracket":
            continue
        target, index = annotation.target
        key, _, band = target.partition(":")
        panel = panels[key]
        if band:
            position = [s.label for s in panel.segments].index(band)
            high = panel.stack(index)[position][2]
            low = panel.stack(index - 1)[position][2]
        else:
            high, low = panel.total(index), panel.total(index - 1)

        x = L["bracket_x"]
        y_high, y_low = scale.y(high), scale.y(low)
        c.rect(x, y_high, L["bracket_weight"], y_low - y_high,
               fill=S.VARIANCE["good"])
        for y, source in ((y_high, index), (y_low, index - 1)):
            x0 = _panel_x(L, panel, source) + L["bar_width"]
            c.line(x0, y, x, y, stroke=S.PAGE["rule"], stroke_width=1)

        change = high - low
        percent = change / low * 100
        text_x = x + L["bracket_weight"] + L["bracket_label_gap"]
        c.text(text_x, (y_high + y_low) / 2 + F["value"] * 0.36,
               f"{change:+,.1f}", size=F["value"])
        width = _text_width(f"{change:+,.1f}", F["value"], False)
        highlight_oval(c, text_x + width + 26, (y_high + y_low) / 2, 24, 11)
        c.text(text_x + width + 26, (y_high + y_low) / 2 + F["value"] * 0.36,
               f"{percent:+,.0f}%", size=F["value"], anchor="middle",
               fill=S.ANNOTATION["highlight"])


def comment_marker(canvas: Canvas, cx: float, cy: float, ref: int) -> None:
    """The numbered circle that ties a band to a note in the margin."""
    blue = S.ANNOTATION["highlight"]
    canvas.ellipse(cx, cy, 9, 9, stroke=blue, stroke_width=1.4, fill="#FFFFFF")
    canvas.text(cx, cy + 4, ref, size=11, fill=blue, anchor="middle")


def _structure_markers(c: Canvas, t: D.Template, L: dict, panels: dict,
                       scale: Scale) -> None:
    for annotation in t.annotations:
        if annotation.kind != "oval" or annotation.ref is None:
            continue
        target, index = annotation.target
        key, _, band = target.partition(":")
        panel = panels[key]
        position = [s.label for s in panel.segments].index(band)
        _, base, top = panel.stack(index)[position]
        cy = (scale.y(top) + scale.y(base)) / 2
        side = L["panels"][panel.key].get("marker_side", "right")
        x = _panel_x(L, panel, index)
        c_x = x + L["bar_width"] if side == "right" else x
        comment_marker(c, c_x, cy, annotation.ref)


# --------------------------------------------------------------------------- #
# C01A layout - measured from template-refs/C01_01A.png
# --------------------------------------------------------------------------- #
#
# One scale for three panels, which is the template's whole argument: 4.57px per
# kEUR everywhere, so 101.9 is 466px tall whether it is drawn as five years of
# business areas, six industries or five regions.

C01A_LAYOUT = dict(
    page=(1280, 720),
    margin_left=20,
    margin_right=20,
    title_rule_y=80,
    message_width=440,          # breaks after "(+12%)", as the original does

    baseline=635.0,
    unit_scale=4.570,          # px per kEUR, shared by all three panels
    bar_width=50.0,
    panels={
        "area": dict(x0=107.0, pitch=72.0, header_dx=50.0, marker_side="right"),
        "industry": dict(x0=939.0, marker_side="left"),
        "region": dict(x0=1051.0, marker_side="left"),
    },
    panel_header_baseline=128.5,
    category_baseline=654.0,
    scenario_baseline=675.0,
    total_gap=6.0,
    legend_gap=9.0,
    label_min_height=14.0,

    divider_x=383.5,
    divider_top=553.0,
    bracket_x=484.0,
    bracket_weight=5.0,
    bracket_label_gap=8.0,

    comments_x=530,
    comments_y=(487, 597),
    comments_width=272.0,       # the next panel, not the page edge
    comment_lead_own_line=True,
    font=dict(entity=13, measure=13.5, period=13, message=13.5, badge=34,
              value=13.5, category=13.5, header=13.5, legend=13.5, comment=12.5,
              footer=9),
)


def render_c01a(template: D.Template = D.C01A, layout: dict | None = None) -> str:
    """Render C01A - Alpha Software net sales, three stacked structure panels."""
    t, L = template, {**C01A_LAYOUT, **(layout or {})}
    c = Canvas(*L["page"], prefix="c01a")
    scale = Scale(L["baseline"], L["unit_scale"])
    panels = {p.key: p for p in t.structure_panels}

    _title_block(c, t, L)
    for panel in t.structure_panels:
        _structure_panel(c, t, L, panel, scale)

    # The axis, drawn under every panel it carries.
    for panel in t.structure_panels:
        spec = L["panels"][panel.key]
        x0 = spec["x0"] - 11
        x1 = (spec["x0"] + spec.get("pitch", 0.0) * (len(panel.categories) - 1)
              + L["bar_width"] + 11)
        c.line(x0, L["baseline"], x1, L["baseline"],
               stroke=S.TEXT["primary"], stroke_width=1.6)

    # Where the actuals stop and the plan starts.
    c.line(L["divider_x"], L["divider_top"], L["divider_x"], L["baseline"],
           stroke=S.TEXT["primary"], stroke_width=1.4)

    _structure_brackets(c, t, L, panels, scale)
    _structure_markers(c, t, L, panels, scale)
    _comments(c, t, L)
    _footer(c, L)
    return c.to_svg()


# --------------------------------------------------------------------------- #
# C02A layout - measured from template-refs/C02_02A.png
# --------------------------------------------------------------------------- #
#
# C01A rotated, and then given a hierarchy. Two things make it its own renderer
# rather than a flag on the last one.
#
# The scale is exact and worth stating: 2.0px per kCHF, so Switzerland's 496
# runs 992px from the axis. Nothing was fitted - the reference is drawn on a
# round number.
#
# And the subtotal rows are not bars. Europe and World are a row of
# swatch-and-figure pairs, one per channel, then the total and its share. That
# is the integrated legend the template requires: the totals row doubles as the
# key, which is why there is no legend box anywhere on the page and why a
# reader never has to match a grey to a caption.

C02A_LAYOUT = dict(
    page=(1280, 720),
    margin_left=20,
    margin_right=20,
    title_rule_y=80,
    message_width=560,
    # The ordering is part of what the chart says, so the subject line states it.
    subject_suffix=" (sorted by total net sales)",

    axis_x=178.5,
    unit_scale=2.0,             # px per kCHF - the reference is on a round number
    row0_top=142.0,
    bar_height=19.0,
    row_pitch=24.0,
    row_gaps={14: 16.0},        # after the Europe subtotal
    label_right=170.0,
    baseline_offset=14.5,
    total_gap=8.0,
    label_min_width=17.0,       # a band narrower than this carries no figure

    subtotal_rule_lift=4.0,
    subtotal_rule_weight=1.8,
    swatch_x=(183.0, 277.0, 370.0),
    swatch_width=12.0,
    swatch_height=11.0,
    swatch_gap=7.0,
    subtotal_total_x=505.0,
    subtotal_share_x=530.0,

    header_baseline=136.0,
    arrow_size=11.0,
    share_row="USA",
    share_dx=58.0,

    comments_x=880,
    comments_y=(497, 555, 613),
    comments_width=300.0,
    comment_lead_own_line=False,
    font=dict(entity=13, measure=13.5, period=13, message=13.5, badge=34,
              value=13.5, category=13.5, header=13.5, legend=13.5, comment=12.5,
              footer=9),
)


def _c02_rows(L: dict, t: D.Template) -> list[float]:
    """Top of every row, with the one gap the hierarchy asks for."""
    tops, y = [], L["row0_top"]
    for i in range(len(t.rows)):
        tops.append(y)
        y += L["row_pitch"] + L["row_gaps"].get(i, 0.0)
    return tops


def render_c02a(template: D.Template = D.C02A, layout: dict | None = None) -> str:
    """Render C02A - Pharmaceutical Inc. net sales by country and channel."""
    t, L = template, {**C02A_LAYOUT, **(layout or {})}
    c = Canvas(*L["page"], prefix="c02a")
    panel = t.structure_panels[0]
    tops = _c02_rows(L, t)
    F = L["font"]

    def x_of(value: float) -> float:
        return L["axis_x"] + value * L["unit_scale"]

    _title_block(c, t, L)

    # Channel headers, centred over the bands of the widest row - which is the
    # first, because the rows are sorted by total.
    for position, (segment, base, top) in enumerate(panel.stack(0)):
        c.text((x_of(base) + x_of(top)) / 2, L["header_baseline"],
               segment.label, size=F["header"], anchor="middle")

    for i, row in enumerate(t.rows):
        y = tops[i]
        mid = y + L["bar_height"] / 2
        baseline = y + L["baseline_offset"]
        bold = "bold" if row.kind == "subtotal" else "normal"
        italic = "italic" if row.label.startswith("Rest of") else "normal"

        c.text(L["label_right"], baseline, row.label, size=F["category"],
               anchor="end", weight=bold, style=italic)

        if row.kind == "subtotal":
            # The integrated legend: a swatch and a figure per channel, then
            # the total and the share it is of the world.
            c.rect(L["axis_x"] - 132, y - L["subtotal_rule_lift"],
                   x_of(panel.total(i)) - L["axis_x"] + 132,
                   L["subtotal_rule_weight"], fill=S.TEXT["primary"])
            for position, segment in enumerate(panel.segments):
                sx = L["swatch_x"][position]
                c.rect(sx, mid - L["swatch_height"] / 2, L["swatch_width"],
                       L["swatch_height"],
                       fill=S.structure_colour(position, len(panel.segments)))
                c.text(sx + L["swatch_width"] + L["swatch_gap"], baseline,
                       f"{segment.at(i):,.0f}".replace(",", " "),
                       size=F["value"], weight="bold")
            total = panel.total(i)
            c.text(L["subtotal_total_x"], baseline,
                   f"{total:,.0f}".replace(",", " "), size=F["value"],
                   weight="bold", anchor="end")
            share = total / panel.total(len(t.rows) - 1) * 100
            highlight_oval(c, L["subtotal_share_x"] + 24, mid, 25, 11)
            c.text(L["subtotal_share_x"] + 24, baseline, f"{share:,.0f}%",
                   size=F["value"], anchor="middle",
                   fill=S.ANNOTATION["highlight"])
            continue

        for position, (segment, base, top) in enumerate(panel.stack(i)):
            if top - base <= 0:
                continue
            fill = S.structure_colour(position, len(panel.segments))
            c.rect(x_of(base), y, x_of(top) - x_of(base), L["bar_height"],
                   fill=fill)
            if x_of(top) - x_of(base) >= L["label_min_width"]:
                c.text((x_of(base) + x_of(top)) / 2, baseline,
                       f"{segment.at(i):,.0f}", size=F["value"],
                       anchor="middle", fill=S.on_fill(fill))

        c.text(x_of(panel.total(i)) + L["total_gap"], baseline,
               f"{panel.total(i):,.0f}".replace(",", " "),
               size=F["value"], style=italic)

        # The one row outside Europe the message singles out carries its share,
        # circled, the way the two subtotal rows carry theirs.
        if row.label == L.get("share_row"):
            share = panel.total(i) / panel.total(len(t.rows) - 1) * 100
            sx = x_of(panel.total(i)) + L["share_dx"]
            highlight_oval(c, sx, mid, 22, 11)
            c.text(sx, baseline, f"{share:,.0f}%", size=F["value"],
                   anchor="middle", fill=S.ANNOTATION["highlight"])

    # The axis every bar starts from.
    c.line(L["axis_x"], L["row0_top"] - 6, L["axis_x"],
           tops[-1] + L["bar_height"], stroke=S.TEXT["primary"], stroke_width=1.6)

    _c02_annotations(c, t, L, panel, tops, x_of)
    _comments(c, t, L)
    _footer(c, L)
    return c.to_svg()


def _c02_annotations(c: Canvas, t: D.Template, L: dict, panel, tops, x_of) -> None:
    """The comment markers, and the arrows that point at what the message names."""
    for annotation in t.annotations:
        target, index = annotation.target
        _, _, band = target.partition(":")
        y = tops[index]
        if annotation.kind == "arrow":
            # Pointing down at the end of a bar the message singles out.
            x = x_of(panel.total(index))
            size = L["arrow_size"]
            c.add(f'<path d="M {x - size / 2:.2f} {y - size - 4:.2f} '
                  f'h {size:.2f} l {-size / 2:.2f} {size:.2f} Z" '
                  f'fill="{S.ANNOTATION["highlight"]}"/>')
            continue
        if annotation.ref is None:
            continue
        position = [s.label for s in panel.segments].index(band)
        segment, base, top = panel.stack(index)[position]
        # Immediately right of the band's own figure, inside the bar - which is
        # where the reference puts it, and the only place on a full row that is
        # not already carrying something.
        centre = (x_of(base) + x_of(top)) / 2
        half = _text_width(f"{segment.at(index):,.0f}", L["font"]["value"], False) / 2
        comment_marker(c, centre + half + 12, y + L["bar_height"] / 2,
                       annotation.ref)

# --------------------------------------------------------------------------- #
# C07C layout - measured from template-refs/C07_07C.png
# --------------------------------------------------------------------------- #
#
# One scale for everything on the page: 0.2339px per kUSD, so a month of 145 and
# a moving annual total of 1 986 are drawn by the same ruler. That is what makes
# the monthly tier and the cumulative lines readable against each other, and it
# is required rather than optional - they share a unit.
#
# The lines sit at the category centres; the columns hang to the left of them,
# plan behind actual, which is the arrangement C05X settled.

C07C_LAYOUT = dict(
    page=(1280, 720),
    margin_left=20,
    margin_right=20,
    title_rule_y=64,
    message_width=520,

    baseline=650.0,
    unit_scale=0.2339,          # px per kUSD - shared by the lines and the tier
    x0=110.0,
    pitch=46.57,
    bar_width=31.0,
    pl_offset=-11.0,            # the plan column, behind and left
    marker=9.0,
    label_gap=7.0,
    divider_after=7,            # the last actual month

    category_baseline=672.0,
    scenario_baseline=690.0,
    series_label_gap=6.0,

    bracket_x=663.0,
    bracket_weight=4.0,
    bracket_label_gap=4.0,
    font=dict(entity=13, measure=13.5, period=13, message=13.5, badge=34,
              value=12.0, category=13.0, series=12.5, footer=9),
)


def render_c07c(template: D.Template = D.C07C, layout: dict | None = None) -> str:
    """Render C07C - Alpha Software net sales, cumulative lines over a monthly tier."""
    t, L = template, {**C07C_LAYOUT, **(layout or {})}
    c = Canvas(*L["page"], prefix="c07c")
    scale = Scale(L["baseline"], L["unit_scale"])
    F = L["font"]

    def x_of(i):
        return L["x0"] + L["pitch"] * i

    _title_block(c, t, L)

    # The monthly tier: plan behind, actual or forecast in front.
    monthly = t.tier("monthly")
    for scenario, offset in (("PL", L["pl_offset"]), ("AC", 0.0), ("FC", 0.0)):
        series = monthly.series_for(scenario)
        if series is None:
            continue
        for i, value in enumerate(series.values):
            if value is None:
                continue
            x = x_of(i) - L["bar_width"] + offset
            scenario_bar(c, scenario, x + L["bar_width"] / 2, L["bar_width"],
                         scale.y(value), scale.y(0))
            if scenario != "PL":
                c.text(x + L["bar_width"] / 2, scale.y(value) - L["label_gap"],
                       f"{value:,.0f}", size=F["value"], anchor="middle")

    # The four lines. Drawn after the tier so a marker sits over a column edge
    # rather than under it.
    def line_points(tier):
        """(index, x, y, scenario, value) for every point the tier defines."""
        out = []
        for series in tier.series:
            for i, value in enumerate(series.values):
                if value is not None:
                    out.append((i, x_of(i), scale.y(value), series.scenario, value))
        out.sort()
        return out

    lines = [t.tier(k) for k in ("cum_pl", "cum_ac", "cum_fc", "mat")]
    drawn = {}
    for tier in lines:
        points = line_points(tier)
        drawn[tier.key] = points
        scenario_line(c, [(p[1], p[2]) for p in points], S.TEXT["primary"], 1.0)

    # The handover. The actual and the forecast are one cumulative told in two
    # scenarios, so the line continues across the join even though the markers
    # change - which is the whole reason the marker carries the scenario.
    if drawn["cum_ac"] and drawn["cum_fc"]:
        a, b = drawn["cum_ac"][-1], drawn["cum_fc"][0]
        scenario_line(c, [(a[1], a[2]), (b[1], b[2])], S.TEXT["primary"], 1.0)

    for tier in lines:
        for i, x, y, scenario, value in drawn[tier.key]:
            scenario_marker(c, scenario, x, y, L["marker"])
            # A label only where the reference prints one. None here means the
            # figure was left off because it would collide, not that the point
            # has no value - so it must suppress the label rather than fall back
            # to the derived number.
            if tier.printed is not None and tier.printed[i] is not None:
                c.text(x, y - L["label_gap"] - 2,
                       f"{tier.printed[i]:,.0f}".replace(",", " "),
                       size=L["font"]["value"], anchor="middle")

    # The series names, at the end of the lines where IBCS puts them.
    for key, name in (("mat", "MAT"),):
        first = drawn[key][0]
        c.text(first[1] - L["series_label_gap"] - 6, first[2] + 4, name,
               size=L["font"]["series"], anchor="end")
    for key in ("cum_fc", "cum_pl"):
        last = drawn[key][-1]
        c.text(last[1] + L["series_label_gap"], last[2] + 4,
               t.tier(key).label, size=L["font"]["series"])

    # The two variances the message names, each bracketing two lines at one
    # month. Drawn from the data, so the sentence is the thing that has to be
    # kept in step rather than the other way round.
    for annotation in t.annotations:
        if annotation.kind != "bracket":
            continue
        _, index = annotation.target
        pair = ("cum_ac", "cum_pl") if index < 8 else ("cum_fc", "cum_pl")
        high = t.tier(pair[0]).series[0].values[index]
        low = t.tier(pair[1]).series[0].values[index]
        y0, y1 = sorted((scale.y(high), scale.y(low)))
        bx = L["bracket_x"]
        colour = S.variance_colour(high - low)
        for y in (y0, y1):
            c.line(x_of(index), y, bx, y, stroke=S.PAGE["rule"], stroke_width=1)
        c.rect(bx, y0, L["bracket_weight"], y1 - y0, fill=colour)
        # The SVG prints what the original printed. Our own arithmetic gives
        # -194 and +83 where the reference shows -192 and +87, because it
        # cumulates unrounded months and prints rounded ones; the bracket is
        # drawn from our data and labelled with theirs, which is the split this
        # project has settled into everywhere else.
        text = annotation.text or f"{high - low:+,.0f}"
        tx = bx + L["bracket_weight"] + L["bracket_label_gap"]
        if annotation.text and annotation.text.startswith("-"):
            highlight_oval(c, tx + 22, (y0 + y1) / 2, 24, 12)
            c.text(tx + 22, (y0 + y1) / 2 + 4, text, size=L["font"]["value"],
                   anchor="middle", fill=S.ANNOTATION["highlight"])
        else:
            c.text(tx, (y0 + y1) / 2 + 4, text, size=L["font"]["value"])

    # Where the actuals stop.
    split = x_of(L["divider_after"]) + L["pitch"] / 2
    c.line(split, 120, split, L["baseline"], stroke=S.PAGE["rule"],
           stroke_width=1)

    c.line(L["x0"] - L["bar_width"] - 14, L["baseline"],
           x_of(len(t.categories) - 1) + 14, L["baseline"],
           stroke=S.TEXT["primary"], stroke_width=1.4)

    previous = None
    for i, (name, scenario) in enumerate(zip(t.categories, t.category_scenarios)):
        x = x_of(i) - L["bar_width"] / 2
        c.text(x, L["category_baseline"], name, size=F["category"],
               anchor="middle")
        if scenario != previous:
            c.text(x, L["scenario_baseline"], scenario, size=F["category"],
                   anchor="middle")
        previous = scenario

    _footer(c, L)
    return c.to_svg()

def movement_fill(scenario: str, positive: bool) -> str:
    """The shade for one bar of a movement tier.

    The polarity is the opposite of a waterfall's, and deliberately so. In a
    statement waterfall the dark bar is the thing being built and the light one
    the deduction from it; in a stock chart the dark bar is the stock going
    *out*. C08H draws an increase at #A6A6A6 over a decrease at #595959 - light
    for what arrives, dark for what leaves - which is the reading a warehouse
    would give it.

    So the palette's two waterfall steps are used with the sign flipped, rather
    than a second pair being invented that would say the same thing twice.
    """
    return S.waterfall_fill(scenario, -1 if positive else +1)


def area_fill(canvas: Canvas, scenario: str) -> str:
    """The paint under a line, for one scenario.

    An area is a fill without an outline to carry the notation, so the scenario
    has to survive being spread over a large pale shape: solid light for an
    actual, lighter still for a plan, and the same 45-degree hatch a forecast
    bar uses. The line's markers say it again at every point, which is what
    makes an area chart readable at all - the fill alone is too weak a signal.
    """
    spec = S.scenario_fill(scenario)
    if spec.hatch:
        return canvas.hatch(spec.hatch, "#FFFFFF")
    # An actual takes the ramp's lightest step; a plan is lighter still but not
    # white - white would read as a hole in the chart rather than as a stock.
    return S.STRUCTURE_RAMP[-1] if scenario == "AC" else "#F2F2F2"


def filled_area(canvas: Canvas, points: Sequence[tuple[float, float]],
                floor: float, paint: str, outline: str | None = None) -> None:
    """A polygon from a run of line points down to the tier floor."""
    if len(points) < 2:
        return
    d = f"M {points[0][0]:.2f} {floor:.2f} "
    d += " ".join(f"L {x:.2f} {y:.2f}" for x, y in points)
    d += f" L {points[-1][0]:.2f} {floor:.2f} Z"
    canvas.add(f'<path d="{d}" fill="{paint}" stroke="none"/>')


# --------------------------------------------------------------------------- #
# C08H layout - measured from template-refs/C08_08H.png
# --------------------------------------------------------------------------- #
#
# Three tiers, one scale: 12.5px per ton everywhere, which is measured rather
# than assumed - nineteen decrease columns give a median of exactly 12.500, and
# the inventory line fits the same figure to within a pixel at nineteen of its
# twenty-one vertices. The two that miss are the hatched forecast quarter, where
# a scan up from the floor stops on a white stripe.
#
# The level is a stock, so it is drawn at quarter *boundaries* - twenty closings
# plus the opening - while the two movement tiers are drawn in the quarters they
# belong to. That is the whole reason the template has three tiers rather than
# two: a flow and a level cannot share an x position and still mean anything.

C08H_LAYOUT = dict(
    page=(1280, 720),
    margin_left=20,
    margin_right=20,
    title_rule_y=80,
    message_width=420,

    x0=35.5,
    quarter=56.0,
    unit_scale=12.5,            # px per ton - shared by all three tiers
    change_zero=157.5,
    level_zero=499.5,
    flow_zero=499.5,
    bar_width=34.0,
    marker=8.0,

    label_gap=6.0,
    caption_x=1166.0,
    caption=(("Inventory change", 168.0), ("Inventory", 268.0),
             ("Increase", 478.0), ("Decrease", 534.0)),
    quarter_baseline=640.0,
    year_baseline=662.0,
    font=dict(entity=13, measure=13.5, period=13, message=13.5, badge=34,
              value=13.0, category=13.0, caption=13.5, footer=9),
)


def render_c08h(template: D.Template = D.C08H, layout: dict | None = None) -> str:
    """Render C08H - Alpha Software raw material, a level over its movements."""
    t, L = template, {**C08H_LAYOUT, **(layout or {})}
    c = Canvas(*L["page"], prefix="c08h")
    F = L["font"]
    n = len(t.categories)

    def bound(i):
        return L["x0"] + L["quarter"] * i

    def centre(i):
        return bound(i) + L["quarter"] / 2

    level = Scale(L["level_zero"], L["unit_scale"])
    change = Scale(L["change_zero"], L["unit_scale"])
    flow = Scale(L["flow_zero"], L["unit_scale"])

    _title_block(c, t, L)

    # The level, drawn first so the movement columns sit over it - which is what
    # the reference does, and what lets an increase be read inside the stock it
    # adds to.
    levels = (D.C08H_OPENING,) + D.C08H_LEVELS
    runs, current, previous = [], [], None
    for i, value in enumerate(levels):
        scenario = t.category_scenarios[min(i, n - 1)] if i else t.category_scenarios[0]
        if i and scenario != previous:
            current.append((bound(i), level.y(value)))
            runs.append((previous, current))
            current = [(bound(i), level.y(value))]
        else:
            current.append((bound(i), level.y(value)))
        previous = scenario
    runs.append((previous, current))
    for scenario, points in runs:
        filled_area(c, points, L["level_zero"], area_fill(c, scenario))
    scenario_line(c, [(bound(i), level.y(v)) for i, v in enumerate(levels)],
                  S.TEXT["primary"], 1.4)

    # The two movement tiers. Increase and decrease share an axis with the
    # level, which is why the increase columns rise into the stock.
    for key, scale, sign in (("change", change, +1), ("flow", flow, +1),
                             ("outflow", flow, -1)):
        tier = t.tier(key)
        for series in tier.series:
            for i, value in enumerate(series.values):
                if value is None or value == 0:
                    continue
                shown = value if key != "change" else value
                top = scale.y(sign * abs(shown) if key != "change" else shown)
                base = scale.y(0)
                positive = (shown > 0) if key == "change" else (sign > 0)
                fill = movement_fill(series.scenario, positive)
                paint, _outline = paint_for(c, series.scenario)
                spec = S.scenario_fill(series.scenario)
                body = paint if spec.hatch else fill
                y0, y1 = sorted((top, base))
                c.rect(centre(i) - L["bar_width"] / 2, y0, L["bar_width"],
                       y1 - y0, fill=body,
                       stroke=spec.outline or S.TEXT["primary"],
                       stroke_width=0.8)
                text = f"{shown:+,.0f}" if key == "change" else f"{abs(shown):,.0f}"
                ty = (y0 - L["label_gap"]) if top <= base else (y1 + L["label_gap"] + 10)
                c.text(centre(i), ty, text, size=F["value"], anchor="middle")

    # The level's own figures, at the boundaries they belong to.
    for i, value in enumerate(levels):
        scenario = t.category_scenarios[min(i, n - 1)]
        scenario_marker(c, scenario, bound(i), level.y(value), L["marker"])
        c.text(bound(i), level.y(value) - L["label_gap"] - 4, f"{value:,.0f}",
               size=F["value"], anchor="middle")

    # The year rules, and the light one where the actuals stop.
    for year in range(5):
        x = bound(year * 4)
        c.line(x, 100, x, L["flow_zero"] + 120, stroke=S.TEXT["primary"],
               stroke_width=1.2)
    split = bound(11)
    c.line(split, 100, split, L["flow_zero"] + 120, stroke=S.PAGE["rule"],
           stroke_width=1)
    for y in (L["change_zero"], L["flow_zero"]):
        c.line(bound(0), y, bound(n), y, stroke=S.TEXT["primary"],
               stroke_width=1.2)

    for i, name in enumerate(t.categories):
        c.text(centre(i), L["quarter_baseline"], name.split()[0],
               size=F["category"], anchor="middle")
        if i % 4 == 0:
            c.text(centre(i), L["year_baseline"], name.split()[1],
                   size=F["category"], anchor="middle")

    for text, y in L["caption"]:
        c.text(L["caption_x"], y, text, size=F["caption"])

    # The level the message names.
    highlight_oval(c, bound(12), level.y(levels[12]), 22, 12)

    _footer(c, L)
    return c.to_svg()


# --------------------------------------------------------------------------- #
# XY primitives - shared by the bubble chart and the scattergram
# --------------------------------------------------------------------------- #


@dataclass(frozen=True)
class XYPlot:
    """The rectangle two continuous axes are drawn in, and the map into it.

    The category charts get their x from a ``CategoryAxis``, which knows how
    many slots there are and hands out centres. An XY chart has no slots: a
    point's position is its value, and the only thing between the two is the
    plot rectangle. Holding the rectangle and both axes in one object is what
    stops a renderer from scaling x against y's range, which is the single
    easiest mistake to make here and draws a picture that looks entirely fine.

    ``bottom`` and ``top`` are the pixel rows of the y axis *minimum* and
    *maximum*, in that order - so the flip lives here once rather than at every
    call site.
    """

    left: float
    right: float
    bottom: float
    top: float
    axis_x: D.Axis
    axis_y: D.Axis

    def x(self, value: float) -> float:
        return self.left + (self.right - self.left) * self.axis_x.fraction(value)

    def y(self, value: float) -> float:
        return self.bottom + (self.top - self.bottom) * self.axis_y.fraction(value)

    def at(self, point: D.Point) -> tuple[float, float]:
        return self.x(point.x), self.y(point.y)


def bubble_radius(size: float, scale: float) -> float:
    """Radius for a bubble of the given value.

    Area proportional to value, so radius goes as its square root. This is the
    whole of the bubble chart's third channel and it is worth stating as its own
    function: making radius proportional instead would overstate a large value
    by the square of the ratio, which on this template would draw VAB's 32
    against V12's 1.1 twenty-nine times too big.
    """
    return scale * sqrt(max(size, 0.0))


def bubble(canvas: Canvas, scenario: str, cx: float, cy: float,
           radius: float) -> None:
    """One bubble, in the fill its scenario takes at this size."""
    colour, opacity = S.bubble_fill(scenario)
    canvas.add(
        f'<circle cx="{cx:.2f}" cy="{cy:.2f}" r="{radius:.2f}" '
        f'fill="{colour}" fill-opacity="{opacity:g}"/>'
    )


def xy_grid(canvas: Canvas, plot: XYPlot, L: dict) -> None:
    """The gridlines, at the tick values the axes declare.

    One weight and one grey throughout, including the two lines that bound the
    plot: nothing here is a zero line. A category chart's baseline is the axis
    the bars stand on and gets emphasis for it; an XY chart's bottom edge is
    just the smallest value the author chose to show.
    """
    grid, width = L["grid"], L["grid_width"]
    for value in plot.axis_y.ticks():
        y = plot.y(value)
        canvas.line(plot.left, y, plot.right, y, stroke=grid, stroke_width=width)
    for value in plot.axis_x.ticks():
        x = plot.x(value)
        canvas.line(x, plot.top, x, plot.bottom, stroke=grid, stroke_width=width)


def xy_axis_labels(canvas: Canvas, plot: XYPlot, L: dict) -> None:
    """Tick labels and both axis titles.

    The axis titles are bold and the tick labels are not, which is the same
    split the title block makes: what the number measures is the subject, the
    number itself is the reading.
    """
    F = L["font"]
    for value in plot.axis_y.ticks():
        canvas.text(L["y_label_right"], plot.y(value) + F["axis"] * 0.36,
                    plot.axis_y.label(value), size=F["axis"], anchor="end")
    for value in plot.axis_x.ticks():
        canvas.text(plot.x(value), L["x_label_baseline"],
                    plot.axis_x.label(value), size=F["axis"], anchor="middle")

    for n, line in enumerate(L["y_title_lines"]):
        canvas.text(L["y_title_right"], L["y_title_baseline"] + n * L["y_title_leading"],
                    line, size=F["axis_title"], weight="bold", anchor="end")
    canvas.text(L["x_title_right"], L["x_label_baseline"], plot.axis_x.title,
                size=F["axis_title"], weight="bold", anchor="end")


def size_legend(canvas: Canvas, t: D.Template, L: dict) -> None:
    """The key for the channel neither axis carries.

    Two lines of heading and one swatch per scenario. The heading is where a
    bubble chart states its third unit - the title block cannot, because three
    measures are plotted and they are in three different units - so it is not
    decoration and it is not optional.

    The swatches are all one size on purpose. They say which scenario a fill
    means, not how big a bubble of a given value would be; sizing them to sample
    values would put a fourth thing on the page for a reader to decode.
    """
    F = L["font"]
    measure, unit = t.size_legend
    # The heading ranges with the swatches, not with their captions: it is the
    # key's title rather than a fourth entry in it.
    canvas.text(L["legend_head_x"], L["legend_head_baseline"], measure,
                size=F["legend"], weight="bold")
    canvas.text(L["legend_head_x"],
                L["legend_head_baseline"] + L["legend_head_leading"],
                f"in {unit}", size=F["legend"])

    for n, (scenario, caption) in enumerate(L["legend_entries"]):
        cy = L["legend_top"] + n * L["legend_pitch"]
        bubble(canvas, scenario, L["legend_x"], cy, L["legend_radius"])
        lines = caption if isinstance(caption, tuple) else (caption,)
        # A two-line caption hangs off the swatch centre rather than sitting on
        # it, so a long entry grows downwards instead of drifting off the key.
        first = (cy + F["legend"] * L["legend_label_drop"]
                 - (len(lines) - 1) * L["legend_leading"] / 2)
        for i, line in enumerate(lines):
            canvas.text(L["legend_text_x"], first + i * L["legend_leading"],
                        line, size=F["legend"])


# --------------------------------------------------------------------------- #
# C10D layout - measured from template-refs/C10_10D.png
# --------------------------------------------------------------------------- #
#
# The plot rectangle runs x 95.5..1127.5 for attractiveness 0.00..1.00 and
# y 643.5..99.5 for share 0.00..1.25 - 1032px and 544px, so the two axes are on
# deliberately different scales and nothing may be carried between them.
#
# The area constant is 15.0px per root-mEUR, recovered by fitting circles to the
# reference's twenty bubbles with centre and radius both free: the regression of
# measured radius on root-value gives a slope of 14.995 and an intercept of
# -1.79px, that intercept being the white outline each bubble carries, which a
# measurement of the visible disc loses at both edges. So 15.0 is the drawn
# figure and the missing 1.8px is the outline, not a fitting residual.

C10D_LAYOUT = dict(
    page=(1280, 720),
    margin_left=11,
    margin_right=11,
    title_rule_y=75.5,
    message_width=310,
    message_leading=19,
    period_baseline=63.0,

    left=95.5, right=1127.5, bottom=643.5, top=99.5,
    area_scale=15.0,
    grid="#A6A6A6",
    grid_width=2.0,

    y_label_right=85.0,
    y_title_right=89.0,
    y_title_baseline=128.0,
    y_title_leading=19.0,
    y_title_lines=("Relative", "market", "share"),
    x_label_baseline=667.0,
    x_title_right=1106.0,

    name_gap=7.0,
    label_drop=0.40,           # value label baseline below a bubble centre

    badge_right=1262.0,
    badge_baseline=51.0,

    legend_x=1161.5,
    legend_text_x=1182.0,
    legend_head_x=1145.0,
    legend_radius=14.0,
    legend_top=156.0,
    legend_pitch=35.25,
    legend_leading=19.0,
    legend_label_drop=0.49,
    legend_head_baseline=115.0,
    legend_head_leading=19.0,
    legend_entries=(("PY", "PY"), ("AC", "AC"),
                    (D.C10D_ACQUISITION, ("Acquisitions", "12/25"))),

    # One size for everything the page says in words, which is what the
    # reference does: solved from the ink width of eight separate strings, all
    # of which land between 15.2 and 15.8px.
    # Two sizes, both solved from ink widths rather than guessed: the title
    # block runs a shade larger than everything inside the plot.
    font=dict(entity=16, measure=16, period=16, message=16, badge=47,
              axis=15.5, axis_title=15.5, legend=15.5, name=15.5, value=15.5,
              footer=9),
)


def render_c10d(template: D.Template = D.C10D, layout: dict | None = None) -> str:
    """Render C10D - Alpha Corp.'s product market portfolio.

    The drawing order is the data order and that is deliberate: prior year,
    then actual, then the two acquisitions. Bubbles overlap and the actual is
    only 80% opaque, so what is painted last is what the eye lands on, and the
    thing the reader is being asked about is where each unit stands *now*.
    """
    t, L = template, {**C10D_LAYOUT, **(layout or {})}
    c = Canvas(*L["page"], prefix="c10d")
    F = L["font"]
    plot = XYPlot(L["left"], L["right"], L["bottom"], L["top"], *t.axes)

    _title_block(c, t, L)
    xy_grid(c, plot, L)

    for point in t.points:
        cx, cy = plot.at(point)
        radius = bubble_radius(point.size, L["area_scale"])
        bubble(c, point.scenario, cx, cy, radius)
        # The name above the mark and the value inside it, both centred on the
        # bubble. Nothing is dodged sideways anywhere in the reference: pairs
        # that look dodged are two bubbles of different sizes with their names
        # over their own centres. Where that would put two names into each
        # other, one of them is simply not drawn.
        if (point.entity, point.scenario) not in D.C10D_UNLABELLED:
            c.text(cx, cy - radius - L["name_gap"], point.entity,
                   size=F["name"], anchor="middle")
        c.text(cx, cy + F["value"] * L["label_drop"], f"{point.size:,.1f}",
               size=F["value"], anchor="middle",
               fill=S.on_fill(S.bubble_fill(point.scenario)[0]))

    xy_axis_labels(c, plot, L)
    size_legend(c, t, L)
    _footer(c, L)
    return c.to_svg()


def scatter_marker(canvas: Canvas, colour: str, cx: float, cy: float,
                   radius: float, opacity: float, outline: str,
                   outline_width: float) -> None:
    """One point of a scattergram.

    Outlined and translucent, both for the same reason: a hundred and fifty
    marks on one page overlap, and a solid opaque disc hides whatever is under
    it entirely. The outline is what keeps two overlapping marks countable, and
    the translucency is what stops the top one erasing the bottom one - between
    them a cluster stays readable as a cluster rather than a blob.
    """
    canvas.add(
        f'<circle cx="{cx:.2f}" cy="{cy:.2f}" r="{radius:.2f}" fill="{colour}" '
        f'fill-opacity="{opacity:g}" stroke="{outline}" '
        f'stroke-width="{outline_width:g}"/>'
    )


def iso_curve(plot: XYPlot, level: float, steps: int = 160
              ) -> list[tuple[float, float]]:
    """The locus of one constant product of the two axes, as pixel points.

    A curve rather than a line, and drawn rather than approximated between its
    endpoints: gross profit is net sales times margin, so equal profit is a
    hyperbola, and a straight line between the same two ends would sit as much
    as a third of the plot away from it in the middle.

    Clipped to the plot on the way out, because the curve runs to infinity at
    a margin of nothing and would otherwise leave the page.
    """
    axis_x, axis_y = plot.axis_x, plot.axis_y
    points: list[tuple[float, float]] = []
    # Walk down from the top of the y axis rather than across x: the curve is
    # near-vertical at the left and near-horizontal at the right, and stepping
    # in x alone draws the steep part as a staircase.
    lo = level * 100.0 / axis_y.maximum if axis_y.maximum else axis_x.maximum
    for i in range(steps + 1):
        margin = lo + (axis_x.maximum - lo) * i / steps
        if margin <= 0:
            continue
        sales = level * 100.0 / margin
        if sales > axis_y.maximum or sales < axis_y.minimum:
            continue
        points.append((plot.x(margin), plot.y(sales)))
    return points


def iso_segment(canvas: Canvas, plot: XYPlot, level: float, fill: str) -> None:
    """The band above one iso-profit curve, filled.

    This is the message made visible. The sentence names a segment - "3 mUSD
    and above" - and shading it is what lets a reader check the claim by
    looking, instead of tracing a curve with a finger.
    """
    curve = iso_curve(plot, level)
    if not curve:
        return
    d = ("M " + " L ".join(f"{x:.2f} {y:.2f}" for x, y in curve)
         + f" L {plot.right:.2f} {plot.top:.2f}"
         + f" L {curve[0][0]:.2f} {plot.top:.2f} Z")
    canvas.add(f'<path d="{d}" fill="{fill}" stroke="none"/>')


def category_legend(canvas: Canvas, L: dict, entries: Sequence[tuple[str, str]],
                    heading: str) -> None:
    """A key for a categorical colour, which a scenario never needs.

    Worth being explicit about, because IBCS's whole point about scenario fills
    is that they need no key - solid dark means measured everywhere, in every
    chart, forever. A product line means nothing outside this page, so it does
    need one, and the key is what marks the colour as carrying a category
    rather than a scenario.
    """
    F = L["font"]
    canvas.text(L["legend_x"], L["legend_head_baseline"], heading,
                size=F["legend"])
    for n, (colour, caption) in enumerate(entries):
        cy = L["legend_top"] + n * L["legend_pitch"]
        scatter_marker(canvas, colour, L["legend_swatch_x"], cy,
                       L["legend_radius"], L["marker_opacity"],
                       L["marker_outline"], L["marker_outline_width"])
        canvas.text(L["legend_text_x"], cy + F["legend"] * 0.36, caption,
                    size=F["legend"])


# --------------------------------------------------------------------------- #
# C09C layout - measured from template-refs/C09_09C.png
# --------------------------------------------------------------------------- #
#
# A square plot on two axes that run 0..35 in different units: x 119.5..1159.5
# for margin and y 635.5..115.5 for net sales, so 29.71px per point of margin
# and 14.86px per mUSD. Nothing may be carried between them - the picture only
# looks square.
#
# The page furniture is C10's: rules at y 75.5 and 704 running x 11..1269, the
# title block at x 11, the badge ending at 1262. The two templates were rendered
# by the same generator and it shows.

C09C_LAYOUT = dict(
    page=(1280, 720),
    margin_left=11,
    margin_right=11,
    title_rule_y=75.5,
    message_width=380,
    message_leading=19,
    period_baseline=63.0,
    badge_right=1262.0,
    badge_baseline=51.0,

    left=119.5, right=1159.5, bottom=635.5, top=115.5,
    grid="#A6A6A6",
    grid_width=2.0,
    segment_fill="#F2F2F2",
    curve="#A6A6A6",
    curve_width=1.6,

    marker_radius=8.7,
    marker_opacity=0.85,
    marker_outline="#404040",
    marker_outline_width=1.2,

    y_label_right=109.0,
    y_title_x=10.0,
    y_title_baseline=108.0,
    y_unit_baseline=127.0,
    x_label_baseline=660.0,
    x_title_right=1264.0,
    x_unit_right=1260.0,
    x_unit_baseline=679.0,

    # The gross-profit captions, which say what the curves are. Placed against
    # the right edge at the height each curve leaves the plot, so the caption is
    # read off the curve rather than matched to it.
    curve_label_x=1165.0,
    curve_caption_x=1177.0,
    curve_caption_baseline=472.0,
    curve_caption_leading=19.0,

    legend_x=1177.0,
    legend_head_baseline=131.0,
    legend_swatch_x=1184.0,
    legend_text_x=1195.0,
    legend_top=143.5,
    legend_pitch=18.0,
    legend_radius=6.5,

    font=dict(entity=16, measure=16, period=16, message=16, badge=47,
              axis=15.5, axis_title=15.5, legend=15.5, name=15.5, footer=9),
)

# Where the reference puts each of the ten names it prints. Transcribed, not
# computed: nine sit a few pixels clear of their marker's top edge and VA-3PO
# is lifted to share a baseline with RX-180 beside it, which is a judgement
# about that corner of the page rather than a rule.
C09C_NAME_OFFSETS = {"VA-3PO": -19.0}
C09C_NAME_DEFAULT = -13.0


def render_c09c(template: D.Template = D.C09C, layout: dict | None = None) -> str:
    """Render C09C - Alpha Corp.'s paper products by margin and net sales.

    Order matters and is the reverse of the bubble chart's. There the last mark
    painted is the one the reader is asked about; here every mark is one of a
    hundred and fifty and none is more important than its neighbour, so the
    lines go down in legend order and the translucency does the rest.
    """
    t, L = template, {**C09C_LAYOUT, **(layout or {})}
    c = Canvas(*L["page"], prefix="c09c")
    F = L["font"]
    plot = XYPlot(L["left"], L["right"], L["bottom"], L["top"], *t.axes)

    _title_block(c, t, L)

    # The segment first, then the curves over it, then the frame, then the
    # points: everything the points are read against is underneath them.
    iso_segment(c, plot, D.C09C_SEGMENT, L["segment_fill"])
    for level in D.C09C_ISO_PROFIT:
        scenario_line(c, iso_curve(plot, level), L["curve"],
                      L["curve_width"])
    for x0, y0, x1, y1 in ((plot.left, plot.top, plot.right, plot.top),
                           (plot.left, plot.bottom, plot.right, plot.bottom),
                           (plot.left, plot.top, plot.left, plot.bottom),
                           (plot.right, plot.top, plot.right, plot.bottom)):
        c.line(x0, y0, x1, y1, stroke=L["grid"], stroke_width=L["grid_width"])

    for line in D.C09C_LINES:
        colour = S.structure_colour(D.C09C_ACCENT[line], accent=True)
        for point in (p for p in t.points if p.group == line):
            cx, cy = plot.at(point)
            scatter_marker(c, colour, cx, cy, L["marker_radius"],
                           L["marker_opacity"], L["marker_outline"],
                           L["marker_outline_width"])

    # The ten the reference names, drawn after every marker so a name is never
    # under a point.
    for point in (p for p in t.points if p.entity):
        cx, cy = plot.at(point)
        drop = C09C_NAME_OFFSETS.get(point.entity, C09C_NAME_DEFAULT)
        c.text(cx, cy + drop, point.entity, size=F["name"], anchor="middle")

    _c09_axes(c, plot, L)
    _c09_curve_captions(c, plot, t, L)
    category_legend(c, L, [(S.structure_colour(D.C09C_ACCENT[n], accent=True), n)
                           for n in D.C09C_LINES], "Product lines")
    _footer(c, L)
    return c.to_svg()


def _c09_axes(canvas: Canvas, plot: XYPlot, L: dict) -> None:
    """Tick labels and the two axis titles, each with its own unit under it."""
    F = L["font"]
    for value in plot.axis_y.ticks():
        canvas.text(L["y_label_right"], plot.y(value) + F["axis"] * 0.36,
                    plot.axis_y.label(value), size=F["axis"], anchor="end")
    for value in plot.axis_x.ticks():
        canvas.text(plot.x(value), L["x_label_baseline"],
                    plot.axis_x.label(value), size=F["axis"], anchor="middle")

    canvas.text(L["y_title_x"], L["y_title_baseline"], plot.axis_y.title,
                size=F["axis_title"], weight="bold")
    canvas.text(L["y_title_x"] + 6, L["y_unit_baseline"],
                f"in {plot.axis_y.unit}", size=F["axis_title"])
    canvas.text(L["x_title_right"], L["x_label_baseline"],
                f"{plot.axis_x.title} in %", size=F["axis_title"],
                weight="bold", anchor="end")
    canvas.text(L["x_unit_right"], L["x_unit_baseline"],
                f"of {plot.axis_x.unit.split('of ', 1)[-1]}",
                size=F["axis_title"], anchor="end")


def _c09_curve_captions(canvas: Canvas, plot: XYPlot, t: D.Template,
                        L: dict) -> None:
    """What the curves are, said once, beside the ends of the curves."""
    F = L["font"]
    canvas.text(L["curve_caption_x"], L["curve_caption_baseline"],
                "Gross profit", size=F["axis_title"], weight="bold")
    canvas.text(L["curve_caption_x"],
                L["curve_caption_baseline"] + L["curve_caption_leading"],
                f"in {t.axes[1].unit}", size=F["axis_title"])
    for level in D.C09C_ISO_PROFIT:
        y = plot.y(level * 100.0 / plot.axis_x.maximum)
        canvas.text(L["curve_label_x"], y + F["axis"] * 0.36, f"{level:.0f}",
                    size=F["axis"])


# --------------------------------------------------------------------------- #
# C11A layout - measured from template-refs/C11_11A.png
# --------------------------------------------------------------------------- #
#
# Six boxes, each 358px wide with a 337px axis inset 10.5px from either edge,
# seven 48px category slots on it. The boxes differ in *height* and in nothing
# else: each is sized to its own range, which is what lets six charts of three
# different units share a page.
#
# The three scales are in the data layer, fitted to 42 measured bar extents
# rather than declared here, because the shared scale is the claim this
# template makes and it should be checked rather than asserted.
#
# Two bar widths: 17px in the ratio columns, 33px in the base-measure column,
# with the first bar centred 16px and 24px past the axis start respectively.
# Nothing follows from the difference - it is the reference's own choice and is
# reproduced.

C11A_LAYOUT = dict(
    page=(1280, 720),
    margin_left=11,
    margin_right=11,
    title_rule_y=51.5,
    # The title block is more compact than C03A's - three lines above a rule at
    # 51 rather than at 80 - because six charts need the room. The baselines
    # are layout keys rather than constants for exactly this reason.
    entity_baseline=15.0,
    subject_baseline=29.0,
    period_baseline=43.0,
    message_x=400.0,
    message_top=16.0,
    message_leading=16.0,
    message_width=290.0,
    badge_right=1262.0,
    badge_baseline=49.0,

    box_stroke="#7F7F7F",
    box_stroke_width=1.0,
    axis_inset=10.5,
    axis_colour="#000000",
    axis_width=1.0,

    subtitle_drop=20.0,        # subtitle baseline below the box top
    value_gap=6.0,             # value baseline above a positive bar's top
    value_drop=15.0,           # value baseline below a negative bar's bottom
    category_drop=17.0,        # year baseline below the zero line
    scenario_drop=31.0,        # AC / PL baseline below the zero line

    split_after=5,             # the PL rule falls before the sixth category
    split_above=64.0,          # ... and runs this far above the zero line
    split_below=27.0,          # ... and this far below it
    split_colour="#000000",
    split_width=1.2,

    # Bar geometry per tree column: (bar width, first bar centre past the axis
    # start). Pitch is 48 everywhere.
    pitch=48.0,
    bars={0: (17.0, 16.0), 1: (17.0, 16.0), 2: (33.0, 24.0)},

    connector="#404040",
    connector_width=1.4,
    operator_radius=14.0,
    operator_fill="#FFFFFF",
    # One spine per gap between tree columns, at the middle of the gap.
    spines={1: 416.5, 2: 837.5},
    # Net sales feeds two links, so its two stubs are offset either side of the
    # box centre rather than laid on top of each other.
    stub_offsets={("net_sales", "ros"): -0.5, ("net_sales", "turnover"): 17.5},

    oval_rx=22.5,
    oval_ry=12.0,
    oval_rise=5.0,             # oval centre above the value label's baseline

    font=dict(entity=13, measure=13.5, period=13, message=13.5, badge=41,
              subtitle=15, value=12.5, category=13, footer=9),
)

# Where each box sits and where its zero line is. What it is counted in is a
# fact about the measure, not about the drawing, so it lives on the TreeNode.
# Transcribed from the reference; the heights differ and the scales do not.
C11A_BOXES = {
    "roi":       dict(x0=27.0, y0=235.0, x1=384.0, y1=555.0, zero=430.0),
    "ros":       dict(x0=448.0, y0=115.0, x1=805.0, y1=382.0, zero=291.0),
    "turnover":  dict(x0=448.0, y0=403.0, x1=805.0, y1=670.0, zero=627.0),
    "return":    dict(x0=869.0, y0=72.0, x1=1227.0, y1=200.0, zero=147.0),
    "net_sales": dict(x0=869.0, y0=242.0, x1=1227.0, y1=455.0, zero=414.0),
    "capital":   dict(x0=869.0, y0=499.0, x1=1227.0, y1=691.0, zero=648.0),
}


def _c11_bar_centre(box: dict, node: D.TreeNode, L: dict, index: int) -> float:
    _, lead = L["bars"][node.column]
    return box["x0"] + L["axis_inset"] + lead + index * L["pitch"]


def _c11_box(canvas: Canvas, t: D.Template, node: D.TreeNode, L: dict) -> None:
    """One boxed chart of the tree.

    Everything here is ordinary column-chart drawing - the interest is in what
    is *not* recomputed: the scale comes from the node's group, so the return
    box and the net sales box beside it are on the same ruler even though one
    peaks at 5.5 and the other at 27.7.
    """
    F = L["font"]
    box = C11A_BOXES[node.key]
    tier = t.tier(node.key)
    scale = Scale(box["zero"], D.C11A_SCALE_PX[node.scale_group])
    width, _ = L["bars"][node.column]

    canvas.rect(box["x0"], box["y0"], box["x1"] - box["x0"],
                box["y1"] - box["y0"], fill="none", stroke=L["box_stroke"],
                stroke_width=L["box_stroke_width"])

    # Measure name bold, unit plain - the same two-run subtitle the page title
    # uses, said again per box because each box is counted in its own unit.
    # Capital turnover prints no unit at all: it is a ratio of two currencies
    # and naming one would be wrong.
    runs = [(tier.label, "bold")]
    if node.unit:
        runs.append((f" in {node.unit}", "normal"))
    canvas.rich_text(box["x0"] + 11, box["y0"] + L["subtitle_drop"], runs,
                     size=F["subtitle"])

    axis_x0 = box["x0"] + L["axis_inset"]
    axis_x1 = box["x1"] - L["axis_inset"]

    for index, scenario in enumerate(t.category_scenarios):
        series = tier.series_for(scenario)
        value = series.values[index] if series else None
        if value is None:
            continue
        cx = _c11_bar_centre(box, node, L, index)
        y_value = scale.y(value)
        if value < 0:
            # A negative column takes the lighter of its scenario's two tones -
            # the same step of the ramp a waterfall gives a subtracting row,
            # and for the same reason: within one box a bar that adds to the
            # measure and one that takes away from it have to be told apart,
            # and they are the same scenario, so the fill cannot carry it
            # twice. Measured off the reference at #7F7F7F against the
            # positive columns' #404040.
            canvas.rect(cx - width / 2, box["zero"], width,
                        y_value - box["zero"],
                        fill=S.waterfall_fill(scenario, -1), stroke=None,
                        stroke_width=0)
        else:
            scenario_bar(canvas, scenario, cx, width, y_value, box["zero"])
        baseline = (y_value - L["value_gap"] if value >= 0
                    else y_value + L["value_drop"])
        canvas.text(cx, baseline, tier.label_for(index, value),
                    size=F["value"], anchor="middle")

    # The zero line over the bars, which is what a negative column is read
    # against, then the plan rule over that.
    # Both rules are drawn on the half pixel. A 1px stroke centred on an
    # integer straddles two rows and each renders at half intensity, which puts
    # a soft two-pixel band where the reference has one crisp line - and, worse,
    # makes a pixel measurement of the result ambiguous by a pixel.
    canvas.line(axis_x0, box["zero"] + 0.5, axis_x1, box["zero"] + 0.5,
                stroke=L["axis_colour"], stroke_width=L["axis_width"])
    split_x = ((_c11_bar_centre(box, node, L, L["split_after"] - 1)
                + _c11_bar_centre(box, node, L, L["split_after"])) / 2)
    canvas.line(split_x + 0.5, box["zero"] - L["split_above"],
                split_x + 0.5, box["zero"] + L["split_below"],
                stroke=L["split_colour"], stroke_width=L["split_width"])

    for index in t.tree.printed_categories:
        canvas.text(_c11_bar_centre(box, node, L, index),
                    box["zero"] + L["category_drop"], t.categories[index],
                    size=F["category"], anchor="middle")
    # The scenario is named once under the first period it applies to, in every
    # box - the integrated legend again, six times.
    for index, name in ((0, "AC"), (L["split_after"], "PL")):
        canvas.text(_c11_bar_centre(box, node, L, index),
                    box["zero"] + L["scenario_drop"], name,
                    size=F["category"], anchor="middle")


def _c11_annotations(canvas: Canvas, t: D.Template, L: dict) -> None:
    """The highlight oval, round the planned ROI the message is about."""
    for a in t.annotations:
        if a.kind != "oval":
            continue
        key, index = a.target
        node = t.tree.node(key)
        box = C11A_BOXES[key]
        scale = Scale(box["zero"], D.C11A_SCALE_PX[node.scale_group])
        value = t.tier(key).series[0].values[index]
        if value is None:
            value = t.tier(key).series[1].values[index]
        cx = _c11_bar_centre(box, node, L, index)
        cy = scale.y(value) - L["value_gap"] - L["oval_rise"]
        highlight_oval(canvas, cx, cy, L["oval_rx"], L["oval_ry"], a.ref)


def _c11_links(canvas: Canvas, t: D.Template, L: dict) -> None:
    """The spines, stubs and operator circles.

    Drawn from the links rather than from a list of coordinates, so a tree that
    gained a node would gain its connector too. Each stub meets a box at that
    box's vertical centre, and each operator circle sits at the centre of the
    box it *produces* - which is what makes the glyph read as belonging to the
    result rather than floating between the two operands.
    """
    colour, weight = L["connector"], L["connector_width"]
    radius = L["operator_radius"]

    def centre_y(key: str) -> float:
        box = C11A_BOXES[key]
        return (box["y0"] + box["y1"]) / 2

    for link in t.tree.links:
        result = t.tree.node(link.result)
        spine_x = L["spines"][result.column + 1]
        operands = (link.left, link.right)
        ys = []
        for operand in operands:
            box = C11A_BOXES[operand]
            y = centre_y(operand) + L["stub_offsets"].get(
                (operand, link.result), 0.0)
            canvas.line(spine_x, y, box["x0"], y, stroke=colour,
                        stroke_width=weight)
            ys.append(y)
        canvas.line(spine_x, min(ys), spine_x, max(ys), stroke=colour,
                    stroke_width=weight)

        # The stub from the circle back to the box the link produces.
        cy = centre_y(link.result)
        canvas.line(C11A_BOXES[link.result]["x1"], cy, spine_x - radius, cy,
                    stroke=colour, stroke_width=weight)
        canvas.add(
            f'<circle cx="{spine_x:.2f}" cy="{cy:.2f}" r="{radius:.2f}" '
            f'fill="{L["operator_fill"]}" stroke="{colour}" '
            f'stroke-width="{weight:g}"/>'
        )
        canvas.text(spine_x, cy + L["font"]["category"] * 0.36, link.glyph,
                    size=L["font"]["category"], anchor="middle")


def render_c11a(template: D.Template = D.C11A, layout: dict | None = None) -> str:
    """Render C11A - Alpha Software Corporation's ROI tree.

    Six charts and the arithmetic between them. The drawing is unremarkable;
    what the template is *for* is the two rules that hold it together - boxes
    in the same unit share a scale, and the connectors state a calculation the
    reader can check against the numbers printed in the boxes.

    The links are drawn last so a spine never runs under a box border, and the
    oval last of all so nothing is drawn over the one value the message names.
    """
    t, L = template, {**C11A_LAYOUT, **(layout or {})}
    c = Canvas(*L["page"], prefix="c11a")

    _title_block(c, t, L)
    for node in t.tree.nodes:
        _c11_box(c, t, node, L)
    _c11_links(c, t, L)
    _c11_annotations(c, t, L)
    _footer(c, L)
    return c.to_svg()


# --------------------------------------------------------------------------- #
# C13D layout - measured from template-refs/C13_13D.png
# --------------------------------------------------------------------------- #
#
# A 4x4 grid of 314x160 panels between rules at x 11/325/640/954/1268 and
# y 59/219/379/539/699. Fifteen cells: Berlin takes two rows, and the bottom
# right is the reference panel the other fourteen are measured against.
#
# **Every panel is on one scale**, and it is the whole argument of the
# template. Fitted rather than declared: regressing 167 measured stem lengths
# on the derived variances gives 0.59268px per percentage point with no
# residual over 0.5px. The one value with no stem is Linz 2025, which is zero.
#
# A pin panel's zero sits 83px below the top of the row it ends in - so
# Berlin's is on the *lower* row's zero line, which is exactly what spanning
# two rows buys it: 243px of headroom for +343% where one row gives 83.
#
# The left inset differs per column - 22, 18, 12, 7 - and the pitch does not.
# Nothing follows from it; it is the reference's own drift and is reproduced.

C13D_LAYOUT = dict(
    page=(1280, 720),
    margin_left=11,
    margin_right=11,
    title_rule_y=51.5,
    entity_baseline=15.0,
    subject_baseline=29.0,
    period_baseline=43.0,
    message_x=400.0,
    message_top=16.0,
    message_leading=16.0,
    message_width=470.0,
    badge_right=1262.0,
    badge_baseline=49.0,

    grid_x=(11, 325, 640, 954, 1268),
    grid_y=(59, 219, 379, 539, 699),
    grid_stroke="#7F7F7F",
    grid_width=1.0,

    leads=(22, 18, 12, 7),
    pitch=24.0,
    zero_drop=83.0,            # below the top of the row a panel ends in
    scale=0.59268,             # px per percentage point, fitted to 167 stems

    stem_width=3.0,
    head_width=8.0,
    head_height=7.0,

    title_x=6.0,
    title_baseline=14.0,
    value_gap=5.0,             # value baseline clear of the head
    label_size_gap=4.0,

    split_after=10,            # the plan rule falls before the eleventh year
    split_top_from_bottom=103.0,
    split_bottom_from_bottom=8.0,
    split_colour="#000000",
    split_width=1.0,

    tick_colour="#D9D9D9",
    tick_width=1.4,
    tick_drop=5.0,
    tick_length=37.0,
    printed_categories=(0, 6, 10),
    # Measured up from the foot of the panel, not down from its zero. The
    # zero moves - Berlin's is a whole row lower than everyone else's - and
    # the labels do not: they sit under the deepest pin the panel can draw,
    # which is where the panel ends.
    year_baseline=22.0,        # above the panel's bottom rule
    scenario_baseline=9.0,

    reference_fill="#ECECEC",
    reference_zero_drop=125.0,
    reference_bar_width=17.0,
    reference_lead=9.0,
    reference_scale=0.5900,

    mark_x=4.0,                # the delta-average mark, in from the panel's right
    mark_columns=(3,),

    oval_rx=17.0,
    oval_ry=10.0,

    font=dict(entity=13, measure=13.5, period=13, message=13.5, badge=41,
              title=11.5, value=10, category=10, mark=11, footer=9),
)


def _c13_slot_x(L: dict, cell: D.PanelCell, index: int) -> float:
    return (L["grid_x"][cell.col] + L["leads"][cell.col]
            + index * L["pitch"])


def _c13_zero(L: dict, cell: D.PanelCell) -> float:
    """A panel's zero line: 83px below the top of the row it ends in."""
    return L["grid_y"][cell.row + cell.row_span - 1] + L["zero_drop"]


def _c13_frame(canvas: Canvas, L: dict) -> None:
    """The grid rules, drawn once rather than as fifteen boxes.

    Every interior rule is shared by the two panels either side of it, so
    drawing panels would draw each of them twice - and Berlin's missing rule,
    the one that would cut it in half, is exactly what makes the grid read as a
    grid with one tall cell in it rather than as sixteen boxes.
    """
    gx, gy = L["grid_x"], L["grid_y"]
    stroke, width = L["grid_stroke"], L["grid_width"]
    for x in gx:
        canvas.line(x + 0.5, gy[0], x + 0.5, gy[-1], stroke=stroke,
                    stroke_width=width)
    for j, y in enumerate(gy):
        for c in range(len(gx) - 1):
            # The rule between Berlin's two rows is the one that is not drawn.
            if j == 1 and c == 3:
                continue
            canvas.line(gx[c], y + 0.5, gx[c + 1], y + 0.5, stroke=stroke,
                        stroke_width=width)


def _c13_ticks(canvas: Canvas, L: dict, cell: D.PanelCell, zero: float) -> None:
    """The grey stubs marking the labelled years.

    Drawn *before* the pins. They sit on the same x as a pin does, and a 1.4px
    grey line laid over a 3px red stem tints it - the stem stops reading as the
    variance colour and a colour-blind reader loses the only channel that says
    which way the variance went. The reference draws the pin over the stub.
    """
    for index in L["printed_categories"]:
        x = _c13_slot_x(L, cell, index)
        canvas.line(x, zero + L["tick_drop"], x,
                    zero + L["tick_drop"] + L["tick_length"],
                    stroke=L["tick_colour"], stroke_width=L["tick_width"])


def _c13_axis_labels(canvas: Canvas, t: D.Template, L: dict,
                     cell: D.PanelCell, zero: float) -> None:
    """The years themselves and the scenario, under one panel.

    The reference prints them under the left-hand column only, and under the
    reference panel. Fourteen panels sharing one category axis do not need it
    stated fourteen times - the grey stubs are what carry it across.
    """
    F = L["font"]
    if cell.col != 0 and cell.key != "average":
        return
    foot = L["grid_y"][cell.row + cell.row_span]
    for index in L["printed_categories"]:
        canvas.text(_c13_slot_x(L, cell, index) + 4,
                    foot - L["year_baseline"],
                    "'" + t.categories[index][2:], size=F["category"])
    for index, name in ((0, "AC"), (L["split_after"], "PL")):
        canvas.text(_c13_slot_x(L, cell, index) + 4,
                    foot - L["scenario_baseline"], name, size=F["category"])


def _c13_split(canvas: Canvas, L: dict, cell: D.PanelCell) -> None:
    """The rule between the measured years and the planned ones."""
    bottom = L["grid_y"][cell.row + cell.row_span]
    x = ((_c13_slot_x(L, cell, L["split_after"] - 1)
          + _c13_slot_x(L, cell, L["split_after"])) / 2)
    canvas.line(x + 0.5, bottom - L["split_top_from_bottom"],
                x + 0.5, bottom - L["split_bottom_from_bottom"],
                stroke=L["split_colour"], stroke_width=L["split_width"])


def _c13_pin_panel(canvas: Canvas, t: D.Template, L: dict,
                   cell: D.PanelCell) -> None:
    """One panel of relative-variance pins."""
    F = L["font"]
    tier = t.tier(cell.key)
    zero = _c13_zero(L, cell)
    canvas.text(L["grid_x"][cell.col] + L["title_x"],
                L["grid_y"][cell.row] + L["title_baseline"],
                cell.label or tier.label, size=F["title"])
    _c13_ticks(canvas, L, cell, zero)

    for index, scenario in enumerate(t.category_scenarios):
        series = tier.series_for(scenario)
        value = series.values[index] if series else None
        if value is None:
            continue
        x = _c13_slot_x(L, cell, index)
        y = zero - value * L["scale"]
        colour = S.variance_colour(value, higher_is_better=True)
        variance_pin(canvas, scenario, x, L["stem_width"], y, zero, colour,
                     head=L["head_width"], head_height=L["head_height"])
        above = value >= 0
        baseline = (y - L["head_height"] / 2 - L["value_gap"] if above
                    else y + L["head_height"] / 2 + L["value_gap"] + F["value"])
        canvas.text(x, baseline, tier.label_for(index, value),
                    size=F["value"], anchor="middle")

    canvas.line(L["grid_x"][cell.col] + 1, zero + 0.5,
                L["grid_x"][cell.col + 1] - 1, zero + 0.5,
                stroke="#000000", stroke_width=1)
    _c13_split(canvas, L, cell)
    _c13_axis_labels(canvas, t, L, cell, zero)


def _c13_reference_panel(canvas: Canvas, t: D.Template, L: dict,
                         cell: D.PanelCell) -> None:
    """The panel every other one is measured against.

    Columns, not pins, and on a shaded ground. Both say the same thing: this is
    the baseline rather than a departure from it. Drawn as pins it would be
    twelve marks on the zero line, because its variance from itself is nothing.
    """
    F = L["font"]
    tier = t.tier(cell.key)
    x0, x1 = L["grid_x"][cell.col], L["grid_x"][cell.col + 1]
    y0, y1 = L["grid_y"][cell.row], L["grid_y"][cell.row + cell.row_span]
    zero = y0 + L["reference_zero_drop"]
    canvas.rect(x0 + 1, y0 + 1, x1 - x0 - 2, y1 - y0 - 2,
                fill=L["reference_fill"], stroke=None, stroke_width=0)
    canvas.text(x0 + L["title_x"], y0 + L["title_baseline"],
                cell.label or tier.label, size=F["title"])
    _c13_ticks(canvas, L, cell, zero)

    for index, scenario in enumerate(t.category_scenarios):
        series = tier.series_for(scenario)
        value = series.values[index] if series else None
        if value is None:
            continue
        x = x0 + L["reference_lead"] + index * L["pitch"] + L["pitch"] / 2
        y = zero - value * L["reference_scale"]
        scenario_bar(canvas, scenario, x, L["reference_bar_width"], y, zero)
        canvas.text(x, y - L["value_gap"], tier.label_for(index, value),
                    size=F["value"], anchor="middle")

    canvas.line(x0 + 1, zero + 0.5, x1 - 1, zero + 0.5,
                stroke="#000000", stroke_width=1)
    _c13_split(canvas, L, cell)
    _c13_axis_labels(canvas, t, L, cell, zero)


def _c13_marks(canvas: Canvas, t: D.Template, L: dict) -> None:
    """The delta-average marks down the right-hand column.

    Said once per column rather than once per panel: every pin on the page is a
    variance from the same reference, and the reference panel beside them is
    the thing itself.
    """
    F = L["font"]
    grid = t.panel_grids["reference"]
    for cell in grid.drawn():
        if cell.col not in L["mark_columns"]:
            continue
        zero = (_c13_zero(L, cell) if cell.kind == "pin"
                else L["grid_y"][cell.row] + L["reference_zero_drop"])
        canvas.text(L["grid_x"][cell.col + 1] - L["mark_x"],
                    zero - 6, "ΔØ" if cell.kind == "pin" else "Ø",
                    size=F["mark"], anchor="end")


def render_c13d(template: D.Template = D.C13D, layout: dict | None = None) -> str:
    """Render C13D - Beta Corp.'s twenty-five locations as small multiples.

    The reference's own arrangement: fourteen panels of pins, Berlin twice the
    height because +343% will not fit in one row, and the location average
    drawn as columns on a shaded ground in the last cell.

    One scale throughout, which is what a small multiple is for. Berlin costs
    the other fourteen most of their range - Salzburg's whole series lives
    inside 93px - and that cost is the point: a variance read in one panel
    means the same in the next.
    """
    t, L = template, {**C13D_LAYOUT, **(layout or {})}
    c = Canvas(*L["page"], prefix="c13d")

    _title_block(c, t, L)
    grid = t.panel_grids["reference"]
    for cell in grid.drawn():
        if cell.kind == "column":
            _c13_reference_panel(c, t, L, cell)
        else:
            _c13_pin_panel(c, t, L, cell)
    _c13_frame(c, L)
    _c13_marks(c, t, L)

    for a in t.annotations:
        if a.kind != "oval":
            continue
        key, index = a.target
        cell = grid.cell(key)
        value = D.C13D_VARIANCE[key][index]
        c_x = _c13_slot_x(L, cell, index)
        c_y = _c13_zero(L, cell) - value * L["scale"]
        highlight_oval(c, c_x, c_y - L["head_height"] / 2 - L["value_gap"]
                       - L["font"]["value"] * 0.35, L["oval_rx"], L["oval_ry"],
                       a.ref)

    _footer(c, L)
    return c.to_svg()


RENDERERS = {"C03A": render_c03a, "C04A": render_c04a,
             "C05X": render_c05x, "C06F": render_c06f,
             "C12A": render_c12a, "T01B": render_t01b,
             "T02A": render_t02a,
             "T03A": render_t03a,
             "T04A": render_t04a,
             "C01A": render_c01a,
             "C02A": render_c02a,
             "C07C": render_c07c,
             "C08H": render_c08h,
             "C09C": render_c09c,
             "C10D": render_c10d,
             "C11A": render_c11a,
             "C13D": render_c13d}


def check_registries() -> list[str]:
    """Every template must appear in every registry that has to know about it.

    Five lists have to stay in step as templates are added - the data layer's
    TEMPLATES, this module's RENDERERS, one of the seven Excel sheet registries,
    a tie-out function, and compare_render's reference image. Adding a template
    to four of them and forgetting the fifth is the easiest mistake in the
    project and the failure is always somewhere unhelpful: a KeyError deep in a
    build, or a template that silently never gets compared to its original.

    C03A is the one allowed exception: its tie-outs are written inline in
    check_ties rather than as a function, because it predates the convention.
    """
    problems = []
    templates = set(D.TEMPLATES)

    if set(RENDERERS) != templates:
        problems.append(f"no SVG renderer for {sorted(templates - set(RENDERERS))}"
                        if templates - set(RENDERERS) else
                        f"renderer with no template: "
                        f"{sorted(set(RENDERERS) - templates)}")

    sheets = (set(L.LAYOUTS) | set(L.TABLE_LAYOUTS) | set(L.STRUCTURE_LAYOUTS)
              | set(L.LINE_LAYOUTS) | set(L.XY_LAYOUTS) | set(L.TREE_LAYOUTS)
              | set(L.PANEL_LAYOUTS))
    if sheets != templates:
        problems.append(f"Excel sheet layouts differ from templates: "
                        f"{sorted(sheets ^ templates)}")

    claimed = [k for k in L.LAYOUTS
               if k in set(L.TABLE_LAYOUTS) | set(L.STRUCTURE_LAYOUTS)
               | set(L.LINE_LAYOUTS) | set(L.XY_LAYOUTS) | set(L.TREE_LAYOUTS)
               | set(L.PANEL_LAYOUTS)]
    if claimed:
        problems.append(f"claimed by two sheet paths: {sorted(claimed)}")

    unchecked = sorted(t for t in templates
                       if t[:3] not in D._CHECKS and not t.startswith("C03"))
    if unchecked:
        problems.append(f"no tie-out function for {unchecked}")

    # Simple layouts restate every per-tier fact, and a missing one fails
    # silently rather than loudly - see check_simple_layouts.
    problems.extend(L.check_simple_layouts())
    return problems



def render(template_id: str) -> str:
    try:
        return RENDERERS[template_id]()
    except KeyError:
        raise ValueError(
            f"no SVG renderer for {template_id!r}; have {', '.join(sorted(RENDERERS))}"
        ) from None


# One registry, in the data layer. A second copy here drifted the moment a
# template was added to one and not the other.
TEMPLATES = D.TEMPLATES

if __name__ == "__main__":
    import argparse
    import sys
    from pathlib import Path

    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description="Render an IBCS template to SVG")
    ap.add_argument("out", nargs="?", type=Path)
    ap.add_argument("--template", default="C03A",
                    help=f"one of {', '.join(sorted(RENDERERS))}")
    args = ap.parse_args()

    out = args.out or P.build_dir() / f"{args.template}.svg"
    out.parent.mkdir(parents=True, exist_ok=True)
    # Nothing renders on a broken registry either: a template missing from one
    # of the five lists fails here rather than somewhere unhelpful later.
    drift = check_registries()
    if drift:
        raise SystemExit("registries out of step:" + chr(10) + "  "
                         + (chr(10) + "  ").join(drift))

    # Nothing renders on a broken tie.
    D.check_ties(TEMPLATES[args.template])
    out.write_text(render(args.template), encoding="utf-8")
    print(f"wrote {out} ({out.stat().st_size:,} bytes)")
