"""Ask whether either renderer still reaches past the template it was handed.

Both engines take a ``Template`` and draw it. Fourteen of the seventeen always
carried everything they needed on that object; three did not, and reached back
into ``ibcs_data`` for module globals instead - C08H for its opening balance,
C09C for its accents and iso-curves, C13D for its panel figures, plus C10D's
suppressed labels and C11A's pixel rulers on the SVG side.

That is a defect on its own terms, but it becomes a dangerous one the moment a
second dataset exists, because **it fails silently**. A global left behind does
not raise; it draws the first dataset's figures on a sheet titled with the
second one's entity. Nobody reviewing the output would see anything wrong,
because nothing looks wrong - the numbers are simply somebody else's.

So this asserts the property rather than the fix: every ``D.<name>`` the two
renderers read must be a *type*, a *helper*, a *registry*, or a ``Template``
itself. Anything else - a tuple of figures, a dict of variances, a bare number -
is data, and data must arrive through the template argument.

Deliberately written against the live module rather than a hard-coded list of
forbidden names. A new constant added to ``ibcs_data`` and read by a renderer
fails here the day it is written, without anyone remembering to update a list.
"""
import ast
import sys
from pathlib import Path

import ibcs_data as D

# The registries. Not figures: a dict of templates keyed by id, and the
# per-template tie-out table. Reading either is asking "which templates exist",
# which is a fair question for a renderer's default argument to ask.
REGISTRIES = {"TEMPLATES", "_CHECKS"}

RENDERERS = ("ibcs_excel.py", "ibcs_svg.py")


def data_reads(path: Path, alias: str = "D") -> list[tuple[int, str]]:
    """Every ``<alias>.<name>`` the file reads, with the line it reads it on."""
    tree = ast.parse(path.read_text(encoding="utf-8"), str(path))
    out = []
    for node in ast.walk(tree):
        if (isinstance(node, ast.Attribute)
                and isinstance(node.value, ast.Name)
                and node.value.id == alias):
            out.append((node.lineno, node.attr))
    return out


def is_leak(name: str, obj: object) -> bool:
    """True where the thing being read is somebody's figures."""
    if name in REGISTRIES:
        return False
    if isinstance(obj, type):                 # a dataclass or an enum
        return False
    if callable(obj):                         # a helper the renderers share
        return False
    if isinstance(obj, D.Template):           # a default argument, not a read
        return False
    return True


def main() -> int:
    here = Path(__file__).parent
    problems: list[str] = []
    checked = 0
    for filename in RENDERERS:
        path = here / filename
        for lineno, name in data_reads(path):
            obj = getattr(D, name, None)
            if obj is None and not hasattr(D, name):
                problems.append(f"{filename}:{lineno} reads D.{name}, "
                                f"which ibcs_data does not define")
                continue
            checked += 1
            if is_leak(name, obj):
                kind = type(obj).__name__
                problems.append(
                    f"{filename}:{lineno} reads D.{name} ({kind}) - that is "
                    f"data, and it must come from the template argument. A "
                    f"second dataset would draw the first one's figures here, "
                    f"without erroring.")

    print(f"checked {checked} references to ibcs_data across "
          f"{len(RENDERERS)} renderers")
    if problems:
        print(f"\n{len(problems)} leak(s):", file=sys.stderr)
        for p in problems:
            print(f"  {p}", file=sys.stderr)
        return 1
    print("PASS: neither renderer reads figures from the data module; "
          "everything drawn arrives on the template it was handed.")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    raise SystemExit(main())
