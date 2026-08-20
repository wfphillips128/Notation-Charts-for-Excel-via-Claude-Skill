"""Build the Progressive workbook from ``ibcs_data_alt``.

Separate from ``ibcs_excel``'s own CLI because that one resolves templates from
``ibcs_data``. Nothing here swaps or monkeypatches anything: ``build`` takes the
templates it is given, which is the whole point of step A.
"""
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

import ibcs_data_alt as A                              # noqa: E402
import ibcs_excel as E                                 # noqa: E402
import test_alt_ties as T                              # noqa: E402


def main() -> int:
    out = Path(sys.argv[1] if len(sys.argv) > 1 else "IBCS-Progressive.xlsx")
    templates = [A.TEMPLATES[n] for n in sorted(A.TEMPLATES)]
    print(f"building {len(templates)} sheet(s): {', '.join(sorted(A.TEMPLATES))}")
    return E.build(templates, out, keep_open=False, export_dir=None,
                   check_ties=T.check_template)


if __name__ == "__main__":
    raise SystemExit(main())
