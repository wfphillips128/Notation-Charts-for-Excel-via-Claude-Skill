"""Does the provenance marking actually mark, and does it stay out of the way?

Two directions, and the second matters as much as the first.

**Engaged.** A template that declares a scenario ``assumed`` must shade that
scenario's typed column differently from the reported one beside it, and the
workbook must carry a Sources sheet saying what was assumed and why. A figure
nobody published, shaded like one that was, is worse than no marking at all -
it launders an assumption into a fact.

**Inert.** A template that declares nothing must come out exactly as it did
before this facility existed: one shade for every typed cell, and no Sources
sheet. The recreation is seventeen such templates, and it is already published.
A feature that quietly restyles shipped work is a defect however good the
feature is.

Needs Excel, like ``test_rescale``: the thing under test is a cell's fill as
Excel actually stored it, and reading that back out of the file is the only
check that cannot agree with a bug in the writer.
"""
from __future__ import annotations

import dataclasses
import sys
import tempfile
from pathlib import Path

try:
    import win32com.client as win32
except ImportError:                                   # pragma: no cover
    win32 = None

import ibcs_data as D
import ibcs_excel as E


def cell_fill(sheet, row: int, col: int) -> int:
    return int(sheet.Cells(row, col).Interior.Color)


def column_of(layout, key: str) -> int:
    """1-based worksheet column for a layout column key."""
    for i, column in enumerate(layout.columns, start=1):
        if column.key == key:
            return i
    raise KeyError(key)


def build_one(template: D.Template, path: Path) -> None:
    problems = E.build([template], path, keep_open=False, export_dir=None)
    if problems:
        raise SystemExit(f"build reported problems: {problems}")


def sheet_names(wb) -> list[str]:
    return [wb.Sheets(i).Name for i in range(1, wb.Sheets.Count + 1)]


def check(results: list[tuple[str, bool, str]], name: str,
          ok: bool, detail: str) -> None:
    results.append((name, ok, detail))
    print(f"  [{'PASS' if ok else 'FAIL'}] {name} - {detail}")


def main() -> int:
    if win32 is None:
        print("error: pywin32 is not installed", file=sys.stderr)
        return 1

    import ibcs_layout as L
    results: list[tuple[str, bool, str]] = []
    plain = D.C03A
    layout = L.layout_for("C03A")
    ref_scenario = layout.reference(plain)
    ref_col = column_of(layout, "ref")
    measure_col = column_of(layout, "measure")

    marked = dataclasses.replace(plain, provenance=(
        D.SourceNote(ref_scenario, "assumed",
                     "Plan is not disclosed; constructed at a 96 combined "
                     "ratio applied to earned premium."),
        D.SourceNote("AC", "filed", "Monthly results release, Exhibit 99."),
    ))

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            for label, template, expect_sources in (
                    ("inert", plain, False),
                    ("engaged", marked, True)):
                out = tmp / f"{label}.xlsx"
                build_one(template, out)
                wb = excel.Workbooks.Open(str(out))
                try:
                    names = sheet_names(wb)
                    check(results, f"{label}: Sources sheet",
                          ("Sources" in names) == expect_sources,
                          f"sheets={names}")

                    sheet = wb.Sheets("C03A")
                    # The first data row of the block; the header rows above it
                    # are not typed inputs and are not shaded as such.
                    row = layout.first_data_row if hasattr(
                        layout, "first_data_row") else 16
                    while row < 60 and cell_fill(sheet, row, measure_col) not in (
                            E.rgb(E.INPUT_FILL), E.rgb(E.ASSUMED_FILL)):
                        row += 1
                    got_ref = cell_fill(sheet, row, ref_col)
                    got_meas = cell_fill(sheet, row, measure_col)

                    want_ref = E.rgb(E.ASSUMED_FILL if expect_sources
                                     else E.INPUT_FILL)
                    check(results, f"{label}: {ref_scenario} column shading",
                          got_ref == want_ref,
                          f"row {row} col {ref_col}: got {got_ref:#08x}, "
                          f"want {want_ref:#08x}")
                    check(results, f"{label}: AC column stays reported",
                          got_meas == E.rgb(E.INPUT_FILL),
                          f"row {row} col {measure_col}: got {got_meas:#08x}")

                    # The footnote: the signal that survives a copy-paste of
                    # values, a photocopy, and a reader who never opens the
                    # Sources tab.
                    last_used = sheet.Cells(sheet.Rows.Count, 1).End(E.XL_UP).Row
                    texts = [str(sheet.Cells(r, 1).Value or "")
                             for r in range(max(1, last_used - 4), last_used + 1)]
                    noted = any("is not reported by the entity" in x
                                for x in texts)
                    check(results, f"{label}: on-sheet note",
                          noted == expect_sources,
                          f"note present={noted} (expected {expect_sources})")
                finally:
                    wb.Close(SaveChanges=False)
        finally:
            excel.Quit()

    failed = [r for r in results if not r[1]]
    print(f"\n{len(results) - len(failed)}/{len(results)} checks passed")
    if failed:
        return 1
    print("PASS: provenance marks what is assumed and leaves the rest alone.")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    raise SystemExit(main())
