# Provenance: saying where a figure came from

A recreation of a published example needs none of this - every figure on the
sheet came from the same picture, and the Read me says so once. A workbook of
somebody's real figures does, because there the scenarios differ *in kind*.
Actuals are reported. Plans are constructed, by you, and a reader has to be
able to see which is which without being told.

This is an **Excel facility**. See *The limit* below before promising it in SVG.

## The type

`SourceNote`, in `ibcs_data.py`, frozen, three fields:

```python
SourceNote(scenario: str,
           basis: Literal["filed", "derived", "assumed"],
           detail: str)
```

`Template.provenance` is a sequence of them and defaults to empty.
`Template.basis_of(scenario)` returns the basis for one scenario, or `""` if the
template never said. Empty is the recreation's answer to everything, which is
why none of this changes a sheet that does not ask for it.

## Three bases, and why two would not do

| Basis | Means | Obligation |
|---|---|---|
| `filed` | The figure appears in a source document | Cite it |
| `derived` | Computed from filed figures by a rule | State the rule |
| `assumed` | A modelling choice nobody published | Say so, in the open |

**The line between filed and assumed is finer than "real or invented".**
Progressive's 96 combined ratio target is *filed* - it is stated in every 10-K.
Applying that target to a month to produce a plan figure is *assumed*. The same
number, a different claim, and only the second one is yours.

`detail` is printed on the sheet, because a basis without its rule is an
assertion. "Assumed" tells a reader to distrust the number; "assumed: the
company's stated 96 combined ratio applied to earned premium" tells them what
to distrust it *for*.

## What the marking changes

All of this is in `ibcs_excel.py`, and all of it is inert on a template that
declares nothing.

**A different hue, not a lighter shade.** `INPUT_FILL` is `#FFF2CC`,
`ASSUMED_FILL` is `#DEEBF7` - deliberately a different colour rather than a
weaker version of the same one, so the distinction survives greyscale printing
and the common colour-vision deficiencies. Those are the two ways a
shade-only signal fails, and a signal that fails silently is worse than none.
`input_fill(template, scenario)` is the single dispatcher; it falls back to
`INPUT_FILL` whenever `basis_of` comes back empty.

**Per cell, not per column.** Eleven reported months and a forecast tail sit in
one column. Shade the column as a block and you paint a constructed figure in
the colour that means reported, which is the one outcome this facility exists
to prevent.

**A variance takes the provenance of the scenario it is measured *against*.**
A gap to a constructed plan is constructed, however solid the actual beside it.
The `ref`, `var_abs` and `var_rel` columns all take
`input_fill(t, layout.reference(t))`. C04A is the case that proves it: it types
the measure and the absolute variance and *derives* the plan, so the plan is
never a typed cell at all - shading only a `ref` column left its whole
constructed side looking reported.

**The sheet footnotes it and the workbook lists it.** `write_provenance_note`
puts a line of text under the data; `write_sources` builds a **Sources** sheet -
Sheet, Scenario, Basis, Detail, with the assumed rows shaded and a legend
defining the three bases. Both, not either: shading does not survive a
copy-paste-values or a photocopy, and a Sources tab only reaches a reader who
thinks to look at another tab. A line under the numbers travels with them.

The Sources sheet is created only when some template in the build declares
provenance, so the recreation's workbooks are unchanged.

## The limit

**Provenance is an Excel facility. `ibcs_svg.py` reads none of it.** No
`SourceNote`, no assumed fill, no footnote, no Sources output. Its only nod to
the subject is a comment beside the two attribution footers.

So never promise a web render that distinguishes a constructed figure from a
reported one. If a chart drawn from someone's own numbers has to carry that
distinction, it has to be the workbook.

## The test, and why it runs both ways

`test_provenance.py`, and it needs Excel - the thing under test is a cell's fill
as Excel actually stored it, and reading that back out of the file is the only
check that cannot agree with a bug in the writer.

- **Engaged.** A template declaring a scenario `assumed` must shade that
  scenario's typed column differently from the reported one beside it, and the
  workbook must carry a Sources sheet.
- **Inert.** A template declaring nothing must come out exactly as it did
  before this facility existed - one shade for every typed cell, no Sources
  sheet. The recreation is seventeen such templates and it is already
  published. **A feature that quietly restyles shipped work is a defect however
  good the feature is.**

## Attribution

**Two disclaimers, because the IBCS® one does nothing about the second
problem.** A real company's figures in polished notation can be read as that
company's own reporting. `read_me_own_data` keeps the middle of the ordinary
Read me, replaces the title and the closing recreation claim, and adds:

- a paragraph explaining the third shade and naming the constructed scenarios;
- the `source_ref` list, so the filings are cited;
- the IBCS® disclaimer, unchanged;
- a non-affiliation statement naming the entity.

Anything built by hand from a real company's figures has to reproduce both.
