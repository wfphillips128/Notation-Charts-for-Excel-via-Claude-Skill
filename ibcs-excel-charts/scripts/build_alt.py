"""Build the Progressive workbook from ``ibcs_data_alt``.

Separate from ``ibcs_excel``'s own CLI because that one resolves templates from
``ibcs_data``. Nothing here swaps or monkeypatches anything: ``build`` takes the
templates it is given, which is the whole point of step A.
"""
import argparse
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

import ibcs_data_alt as A                              # noqa: E402
import ibcs_excel as E                                 # noqa: E402
import ibcs_layout as L                                # noqa: E402
import test_alt_ties as T                              # noqa: E402


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("out", nargs="?", default="IBCS-Progressive.xlsx")
    parser.add_argument("--simple", action="store_true",
                        help="the base-tier version of the templates that "
                             "have a useful one")
    parser.add_argument("--template",
                        help="comma-separated ids; the default is every "
                             "template the dataset carries, or the simple "
                             "ones with --simple")
    parser.add_argument("--doc", help="write the by-hand build guide here")
    parser.add_argument("--export-dir", type=Path, default=None,
                        help="write a PNG of each sheet's charts here. The "
                             "panel grid has no gate that can see a label "
                             "sitting on a rule, so looking at the picture is "
                             "the check - and until now this build could not "
                             "produce one.")
    args = parser.parse_args(argv[1:])

    # Which templates have a simple form is the library's fact, not a list
    # retyped here. Two sources, because there are two ways to have one: a
    # tier stack drops tiers and needs a smaller SheetLayout, while a
    # structure chart is simplified in its *data* - `D.simplify` drops the
    # extra panels - and reuses the layout it already has.
    #
    # The rest have none: a scattergram has no tiers to drop, a line chart's
    # tiers are all measures, and a table reduced to one column block is a
    # list rather than a report.
    simple_ids = sorted(set(L.SIMPLE_LAYOUTS) | {
        name for name in A.TEMPLATES if L.is_structure(name)})

    if args.template:
        names = [n.strip() for n in args.template.split(",") if n.strip()]
    elif args.simple:
        names = [n for n in simple_ids if n in A.TEMPLATES]
    else:
        names = sorted(A.TEMPLATES)

    unknown = [n for n in names if n not in A.TEMPLATES]
    if unknown:
        print(f"error: {', '.join(unknown)} not in this dataset - it carries "
              f"{', '.join(sorted(A.TEMPLATES))}", file=sys.stderr)
        return 1
    if args.simple:
        without = [n for n in names if n not in simple_ids]
        if without:
            print(f"error: {', '.join(without)} have no simple variant; the "
                  f"ones that do are {', '.join(simple_ids)}", file=sys.stderr)
            return 1

    templates = [A.TEMPLATES[n] for n in names]
    print(f"building {len(templates)} "
          f"{'simple ' if args.simple else ''}sheet(s): {', '.join(names)}")
    return E.build(templates, Path(args.out), keep_open=False,
                   export_dir=args.export_dir,
                   simple=args.simple,
                   doc=Path(args.doc) if args.doc else None,
                   check_ties=T.check_template)


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
