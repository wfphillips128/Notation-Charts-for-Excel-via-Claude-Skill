"""Sample the IBCS Institute's own template renderings for their exact fill colours.

IBCS deliberately prescribes no colour codes (v2.0, rule UN 4.1, p52: "Although
IBCS does not prescribe exact codes for these colors, they should still be
consistent within an organization"). What the standard binds is the *semantic*
system - solid dark for measured, solid light for earlier measured, outlined for
fictitious, hatched for expected, and variance colour by impact rather than sign.

So there is no canonical palette to recover. What there is, is the palette the
IBCS Institute uses in its own published template images. Sampling those gives us
a house palette that reproduces the originals pixel-for-pixel, which is what the
proof workbook needs. It is documented as a house choice, not as a rule.

This is a one-off provenance tool: run it, eyeball the report, and hand-map the
significant colours to semantic roles in ibcs-palette.json. It is kept in the
repo so the palette's origin stays auditable.

Usage:
    python extract_palette.py                 # report on the default reference set
    python extract_palette.py C03_03A.png     # report on specific images
"""

from __future__ import annotations

import sys
from collections import Counter
from pathlib import Path

from PIL import Image

import ibcs_paths as P

# Not distributed - see ibcs_paths. Point IBCS_TEMPLATE_REFS at your copy.
REFS = P.refs_dir()

# Between them these five carry every fill in the notation system: scenario
# greys and FC hatch (C03A), variance green/red at both column and pin scale
# (C03A, C13D, T04A), the categorical accent colours (C09C), and the annotation
# blue used for highlight ovals and outlier arrows (C01A).
DEFAULT_SET = [
    "C03_03A.png",
    "T04_T04A.png",
    "C13_13D.png",
    "C01_01A.png",
    "C09_09C.png",
]

# A flat fill covers thousands of pixels; antialiasing on a glyph edge covers a
# handful. This threshold is the whole filter - anything rarer is edge noise.
MIN_PIXELS = 800


def sample(path: Path) -> list[tuple[tuple[int, int, int], int, tuple[int, int]]]:
    """Return [(rgb, pixel_count, first_seen_xy)] for flat fills in one image."""
    img = Image.open(path).convert("RGB")
    width, height = img.size
    pixels = img.load()

    counts: Counter[tuple[int, int, int]] = Counter()
    first_seen: dict[tuple[int, int, int], tuple[int, int]] = {}

    for y in range(height):
        for x in range(width):
            rgb = pixels[x, y]
            counts[rgb] += 1
            if rgb not in first_seen:
                first_seen[rgb] = (x, y)

    return [
        (rgb, n, first_seen[rgb])
        for rgb, n in counts.most_common()
        if n >= MIN_PIXELS
    ]


def hexcode(rgb: tuple[int, int, int]) -> str:
    return "#{:02X}{:02X}{:02X}".format(*rgb)


def describe(rgb: tuple[int, int, int]) -> str:
    """Rough bucket, purely to make the report scannable."""
    r, g, b = rgb
    spread = max(rgb) - min(rgb)
    if spread < 12:
        if r > 240:
            return "white / near-white"
        if r < 40:
            return "black / near-black"
        return f"grey {round(r / 255 * 100)}%"
    if g > r and g > b:
        return "green"
    if r > g and r > b:
        return "red / orange"
    if b > r and b > g:
        return "blue"
    return "mixed"


def main(argv: list[str]) -> int:
    names = argv[1:] or DEFAULT_SET
    for name in names:
        path = REFS / name if not Path(name).is_absolute() else Path(name)
        if not path.exists():
            print(f"!! missing: {path}")
            continue
        print(f"\n=== {path.name} " + "=" * (56 - len(path.name)))
        print(f"{'hex':<9} {'pixels':>8}  {'at (x,y)':<14} bucket")
        for rgb, n, xy in sample(path):
            print(f"{hexcode(rgb):<9} {n:>8}  {str(xy):<14} {describe(rgb)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
