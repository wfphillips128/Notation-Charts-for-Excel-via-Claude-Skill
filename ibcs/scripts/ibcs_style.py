"""Semantic style constants for IBCS notation, loaded from assets/ibcs-palette.json.

Engine-agnostic on purpose: this module knows nothing about SVG or Excel, so the
two renderers cannot drift apart on what an AC bar looks like. Both import from
here; neither hard-codes a colour.

The important thing this module encodes is that IBCS notation is *semantic*.
Callers ask for "the fill for a forecast" or "the colour for this variance given
that lower is better", never for a hex code. Getting that indirection right is
what makes the notation survive a palette change.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Literal

PALETTE_PATH = Path(__file__).resolve().parent.parent / "assets" / "ibcs-palette.json"

with PALETTE_PATH.open(encoding="utf-8") as fh:
    PALETTE = json.load(fh)


# --------------------------------------------------------------------------- #
# Scenarios
# --------------------------------------------------------------------------- #

Scenario = Literal["AC", "PY", "PL", "BU", "FC"]

SCENARIOS: tuple[Scenario, ...] = ("AC", "PY", "PL", "BU", "FC")


@dataclass(frozen=True)
class Hatch:
    """Diagonal stripe spec for the FC (expected) fill."""

    colour: str
    angle_deg: int
    direction: str
    pitch_px: float
    stroke_px: float

    def scaled(self, width_px: float) -> "Hatch":
        """Rescale the stripe geometry to a render width other than the reference."""
        factor = width_px / PALETTE["meta"]["reference_render_size"][0]
        return Hatch(
            colour=self.colour,
            angle_deg=self.angle_deg,
            direction=self.direction,
            pitch_px=self.pitch_px * factor,
            stroke_px=self.stroke_px * factor,
        )


@dataclass(frozen=True)
class Fill:
    """How a visualisation element for one scenario is filled and outlined."""

    scenario: Scenario
    fill: str
    outline: str | None
    hatch: Hatch | None

    @property
    def is_fictitious(self) -> bool:
        """PL and BU are outlined because they describe numbers nobody measured."""
        return self.scenario in ("PL", "BU")

    @property
    def is_expected(self) -> bool:
        return self.scenario == "FC"


def _build_fill(scenario: Scenario) -> Fill:
    spec = PALETTE["scenario"][scenario]
    hatch_spec = spec.get("hatch")
    hatch = (
        Hatch(
            colour=hatch_spec["colour"],
            angle_deg=hatch_spec["angle_deg"],
            direction=hatch_spec["direction"],
            pitch_px=hatch_spec["pitch_px"],
            stroke_px=hatch_spec["stroke_px"],
        )
        if hatch_spec
        else None
    )
    return Fill(scenario=scenario, fill=spec["fill"], outline=spec.get("outline"), hatch=hatch)


FILLS: dict[str, Fill] = {s: _build_fill(s) for s in SCENARIOS}


def scenario_fill(scenario: Scenario) -> Fill:
    """The fill spec for a scenario. Raises rather than guessing on a typo."""
    try:
        return FILLS[scenario]
    except KeyError:
        raise ValueError(
            f"unknown scenario {scenario!r}; expected one of {', '.join(SCENARIOS)}"
        ) from None


# --------------------------------------------------------------------------- #
# Variances
# --------------------------------------------------------------------------- #

VARIANCE = PALETTE["variance"]


def variance_colour(
    value: float,
    higher_is_better: bool = True,
    *,
    neutral: bool = False,
    greyscale: bool = False,
) -> str:
    """Colour a variance by its IMPACT, which is the rule people get wrong.

    A cost that came in 40 over plan is a positive number and an undesirable
    outcome, so it is red. Passing ``higher_is_better=False`` for cost and expense
    measures is what makes that happen; the sign of ``value`` alone never decides.

    ``neutral=True`` is for variances where desirability is genuinely ambiguous -
    headcount, inventory, a mix shift - which IBCS colours blue rather than
    forcing into a good/bad frame.
    """
    palette = VARIANCE["greyscale_fallback"] if greyscale else VARIANCE

    if neutral or value == 0:
        return palette["neutral"]

    desirable = (value > 0) if higher_is_better else (value < 0)
    return palette["good"] if desirable else palette["bad"]


WATERFALL = PALETTE["waterfall"]


def waterfall_fill(scenario: Scenario, sign: int) -> str:
    """The fill for one segment of a waterfall panel.

    A waterfall asks something of the notation that a bar chart does not: within
    one panel, a line that adds to the running total and a line that subtracts
    from it have to be distinguishable, and both are the same scenario. One
    scenario fill cannot say two things, so each panel takes two steps of the
    grey ramp and the panel header carries the scenario.

    ``sign`` is the row's, +1 adds and -1 subtracts - the same number that drives
    the running total and the impact colour, so the three cannot disagree.
    """
    try:
        pair = WATERFALL[scenario]
    except KeyError:
        raise ValueError(
            f"no waterfall shades for {scenario!r}; the palette defines "
            f"{', '.join(k for k in WATERFALL if k in SCENARIOS)}"
        ) from None
    return pair["adds"] if sign > 0 else pair["subtracts"]


BUBBLE = PALETTE["bubble"]


def bubble_fill(scenario: Scenario) -> tuple[str, float]:
    """The fill and opacity for one bubble, as (colour, opacity).

    A bubble is a filled area rather than a bar, and the scenario fills are
    calibrated for bars. Two things change and both are measured: the actual is
    drawn at 80% opacity so it does not hide the prior year it overlaps, and the
    prior year takes a lighter grey because a mark this size carries far more
    weight than a bar does at the same tone.

    A class that is not a scenario at all takes the highlight blue. A bubble
    chart may carry one - C10's two December acquisitions, which have no prior
    year to be compared with and are the subject of the message - and that is
    precisely what the annotation channel is for: it can never be mistaken for
    a scenario grey, because it is not a grey.
    """
    spec = BUBBLE.get(scenario)
    if spec is not None:
        return spec["fill"], float(spec["opacity"])
    if scenario in SCENARIOS:
        return scenario_fill(scenario).colour, 1.0
    return ANNOTATION["highlight"], 1.0


# --------------------------------------------------------------------------- #
# Axes
# --------------------------------------------------------------------------- #

AXIS = PALETTE["axis"]


def reference_axis(scenario: Scenario) -> dict:
    """Style for the axis of a variance tier.

    The variance axis carries the reference scenario - solid light for a PY
    comparison, a double rule for a PL or BU comparison. That is why a correctly
    drawn variance tier needs no legend: the axis already says what the bars are
    measured against.
    """
    if scenario == "PY":
        return dict(AXIS["reference_PY"])
    if scenario in ("PL", "BU"):
        return dict(AXIS["reference_PL"])
    raise ValueError(
        f"{scenario!r} cannot be a variance reference; use PY, PL or BU"
    )


# --------------------------------------------------------------------------- #
# Everything else
# --------------------------------------------------------------------------- #

STRUCTURE_RAMP: list[str] = PALETTE["structure"]["ramp"]
ACCENT_RAMP: list[str] = PALETTE["structure"]["accent"]["ramp"]
ANNOTATION = PALETTE["annotation"]
TEXT = PALETTE["text"]
PAGE = PALETTE["page"]


def on_fill(colour: str) -> str:
    """Black or white, whichever can be read on this fill.

    A stacked structure chart puts its figures *inside* the bands, so the label
    colour is decided by the band rather than chosen once. Asked here rather
    than worked out in each renderer, so a band cannot come out with a readable
    label on the page and an unreadable one in the workbook.
    """
    r, g, b = (int(colour.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4))
    return TEXT["primary"] if (0.299 * r + 0.587 * g + 0.114 * b) > 140 else "#FFFFFF"


def structure_colour(index: int, count: int | None = None, accent: bool = False) -> str:
    """Pick a category colour from the grey ramp (or the accent ramp).

    When ``count`` is given and is smaller than the ramp, the ramp is sampled
    across its full range rather than taken from the top, so a three-segment
    stack uses dark / mid / light instead of three near-identical darks.
    """
    ramp = ACCENT_RAMP if accent else STRUCTURE_RAMP
    if count is None or count >= len(ramp) or count <= 1:
        return ramp[index % len(ramp)]
    step = (len(ramp) - 1) / (count - 1)
    return ramp[round(index * step)]


# Geometry, expressed as ratios so both renderers scale identically. These are
# measured off the reference renders rather than taken from the rule text, which
# names the elements ("pins for relative variances, columns for absolute") but
# fixes no proportions. In C03A at 1280px wide: measure bar 38px, pin stem 5px,
# head marker 8px square.
PIN_WIDTH_RATIO = 5 / 38      # pin stem width, as a fraction of a measure bar
PIN_HEAD_RATIO = 8 / 5        # head square edge, as a multiple of the stem width
LABEL_GAP_RATIO = 0.35        # gap between element end and its label, in bar widths
TIER_GAP_RATIO = 0.18         # vertical gap between stacked tiers, in tier heights


__all__ = [
    "PALETTE",
    "Scenario",
    "SCENARIOS",
    "Fill",
    "Hatch",
    "FILLS",
    "scenario_fill",
    "variance_colour",
    "waterfall_fill",
    "WATERFALL",
    "reference_axis",
    "structure_colour",
    "on_fill",
    "STRUCTURE_RAMP",
    "ACCENT_RAMP",
    "ANNOTATION",
    "TEXT",
    "PAGE",
    "PIN_WIDTH_RATIO",
    "PIN_HEAD_RATIO",
    "LABEL_GAP_RATIO",
    "TIER_GAP_RATIO",
]
