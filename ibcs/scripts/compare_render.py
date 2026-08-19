"""Rasterise a rendered SVG and stack it against the IBCS original for diffing.

The workflow this supports: render, look, adjust, repeat. Producing a single
image with our version above the original (rather than two files) is what makes
a misplaced tier or a wrong bar width obvious at a glance instead of requiring
pixel arithmetic.

Usage:
    python compare_render.py C03A
"""

from __future__ import annotations

import sys
from pathlib import Path

import resvg_py
from PIL import Image, ImageDraw

import ibcs_paths as P

BUILD = P.build_dir()
# The IBCS(R) reference renders, which are not distributed with this
# project - see ibcs_paths. This script is one of only two that need them.
REFS = P.refs_dir()

REFERENCE = {
    "C03A": "C03_03A.png", "C04A": "C04_04A.png",
             "C05X": "C05_05X.png",
             "C06F": "C06_06F.png",
             "C12A": "C12_12A-1.png",
             "T01B": "T01_T01B.png",
             "T02A": "T02_T02A.png",
             "T03A": "T03_T03A.png",
             "T04A": "T04_T04A.png",
             "C01A": "C01_01A.png",
             "C02A": "C02_02A.png",
             "C07C": "C07_07C.png",
             "C08H": "C08_08H.png",
             "C09C": "C09_09C.png",
             "C10D": "C10_10D.png",
             "C11A": "C11_11A.png",
             "C13D": "C13_13D.png"}


def rasterise(svg_path: Path, png_path: Path) -> Image.Image:
    svg = svg_path.read_text(encoding="utf-8")
    raw = resvg_py.svg_to_bytes(svg_string=svg)
    png_path.write_bytes(bytes(raw))
    return Image.open(png_path).convert("RGB")


def stack(ours: Image.Image, theirs: Image.Image, out: Path) -> None:
    if ours.size != theirs.size:
        ours = ours.resize(theirs.size, Image.LANCZOS)
    w, h = theirs.size
    band = 22
    canvas = Image.new("RGB", (w, h * 2 + band * 2), "#FFFFFF")
    draw = ImageDraw.Draw(canvas)
    draw.text((8, 6), "OURS", fill="#0064FF")
    canvas.paste(ours, (0, band))
    draw.text((8, h + band + 6), "IBCS ORIGINAL", fill="#0064FF")
    canvas.paste(theirs, (0, h + band * 2))
    canvas.save(out)


def grid(templates: list[str], columns: int = 4) -> Path:
    """Tile the rendered templates at the size a website grid would show them.

    Worth doing early rather than at handover: a chart that is unreadable at a
    quarter width is a design finding, and the only way to see it is to look at
    it small.
    """
    tiles = []
    for name in templates:
        svg = BUILD / f"{name}.svg"
        if svg.exists():
            tiles.append(rasterise(svg, BUILD / f"{name}.png"))
    if not tiles:
        raise SystemExit("no rendered templates to tile")

    cell_w = 1600 // columns
    cell_h = int(cell_w * tiles[0].height / tiles[0].width)
    rows = (len(tiles) + columns - 1) // columns
    canvas = Image.new("RGB", (cell_w * columns, cell_h * rows), "#FFFFFF")
    for i, tile in enumerate(tiles):
        canvas.paste(tile.resize((cell_w, cell_h), Image.LANCZOS),
                     ((i % columns) * cell_w, (i // columns) * cell_h))
    out = BUILD / "grid.png"
    canvas.save(out)
    return out


def collision_page(templates: list[str]) -> Path:
    """Inline several charts into one document, to catch clashing SVG ids.

    Two charts that each define a pattern called the same thing will quietly
    share whichever one the browser saw last. Rendering them together is the only
    way to see it - separately, both look perfect.
    """
    parts = []
    for name in templates:
        svg = BUILD / f"{name}.svg"
        if svg.exists():
            parts.append(f"<figure><figcaption>{name}</figcaption>"
                         f"{svg.read_text(encoding='utf-8')}</figure>")
    out = BUILD / "inline_collision_check.html"
    out.write_text(
        "<!doctype html><meta charset='utf-8'>"
        "<title>IBCS inline collision check</title>"
        "<style>body{font:14px system-ui;margin:20px}"
        "figure{margin:0 0 24px}svg{width:100%;height:auto;border:1px solid #ddd}"
        "</style>" + "".join(parts),
        encoding="utf-8")
    return out


def main(argv: list[str]) -> int:
    args = argv[1:]
    if args and args[0] == "--grid":
        names = args[1:] or sorted(REFERENCE)
        print(f"wrote {grid(names)}")
        print(f"wrote {collision_page(names)}")
        return 0

    template = args[0] if args else "C03A"
    svg_path = BUILD / f"{template}.svg"
    if not svg_path.exists():
        print(f"!! {svg_path} not found - render it first")
        return 1

    ours = rasterise(svg_path, BUILD / f"{template}.png")
    theirs = Image.open(REFS / REFERENCE[template]).convert("RGB")
    out = BUILD / f"{template}_compare.png"
    stack(ours, theirs, out)
    print(f"ours   {ours.size}\ntheirs {theirs.size}\nwrote  {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
