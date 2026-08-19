# IBCS Charts - Simple Variants Only - how it is built

Every formula quoted here was **read back out of `IBCS Charts - Simple Variants Only.xlsx`** after it was written, so this document and that workbook cannot disagree. Where it says a formula sits in `C03A!E13`, it does.

This is the *simple* workbook: each template drawn with its base tier only. The complex workbook is the same construction with every tier present, and the techniques below are identical.

| | |
|---|---|
| Sheets | 8 (a Read me and 7 templates) |
| Formula cells | 3060 |
| Typed cells | 245 |
| Chart objects | 8 |
| Distinct formula shapes | 170 |

> **One note on what you will see if you open the XML.** Excel stores functions added after 2007 with an `_xlfn.` prefix - `_xlfn.AGGREGATE`, `_xlfn.UNICHAR`. That is a storage detail; you type them without it. This document strips the prefix everywhere.

---

## What every sheet has in common

Ten techniques carry the whole workbook. Read these once and the sheet-by-sheet section below is mostly addresses.

### 1. Six typed lines, and every word derived from them

Rows 1 to 6 of every sheet are the only text anybody types: entity, measure, unit, period, reference scenario, message. The subject line the reader sees is a formula over them:

```
=B2&" in "&B3
```

Read from `C03A!B7`. Retype the unit and the subject line follows.

Tier captions are formulas over the same cells, so a caption cannot describe a comparison the sheet is not making:

*(no example of this in this workbook)*

On a chart sheet the caption *cell* is then linked to a text box on the chart - two hops, one source. **Nothing a chart says is typed into the chart.** Type it in a cell and link the box to the cell, every time.

### 2. Two zones, and nothing straddles them

Each sheet is a data zone of declared width on the left and the charts to its right, with the chart zone's left edge measured from where the data ends. The print area is the chart zone alone:

| Sheet | Cells used | Print area |
|---|---|---|
| C01A | `A1:E53` | `'C01A'!$G$1:$L$71` |
| C02A | `A1:T47` | `'C02A'!$W$1:$AG$73` |
| C03A | `A1:AC22` | `'C03A'!$AD$1:$AO$21` |
| C04A | `A1:Y29` | `'C04A'!$Z$1:$AI$32` |
| C05X | `A1:AU27` | `'C05X'!$AV$1:$BI$52` |
| C06F | `A1:AM30` | `'C06F'!$AN$1:$BC$35` |
| C12A | `A1:AU30` | `'C12A'!$AV$1:$BD$33` |

**Never `AutoFit` a data block a chart is positioned against.** One long string moves the boundary and the charts end up drawn over the source data. Column widths here are declared, never computed from content.

### 3. Splitting a series by what it means, not by its sign

A variance is drawn in two colours, and Excel gives one colour to a series. So the split happens in the *cells*: one column per direction, each `NA()` where the other owns the point.

```
=IF($E11>0,$E11,NA())
```

Read from `C03A!I11`.

`NA()` and not `""` - an empty string plots as a **zero**, which puts a bar of nothing on the axis where there should be no bar at all. That one substitution is the most common way to get a chart that looks broken for no visible reason.

The same split does the work for impact: within one variance column of a statement the favourable values sit on **both** sides of the axis, because a cost line up is adverse and a cost line down is favourable. Split by impact, never by sign, or every cost overrun comes out green.

### 4. The scale block - why the charts follow your numbers

This is the one technique with no equivalent in a hand-drawn chart, and it exists because **Excel will not bind an axis bound to a formula**. An axis maximum is a number you type; it cannot be `=MAX(...)`. So a chart built for one set of figures clips or shrinks the moment somebody pastes their own over the top.

The way round it is to invert the problem: instead of scaling the axis to the data, **scale the data to a fixed axis**. Three parts:

**A span cell** - one per group of tiers sharing a unit - holding the largest magnitude anywhere in that group:

```
=MAX(AGGREGATE(4,6,C11:C22,D11:D22,E11:E22,I11:I22,J11:J22),-AGGREGATE(5,6,C11:C22,D11:D22,E11:E22,I11:I22,J11:J22),0.000000001)
```

Read from `C03A!O11`. `AGGREGATE(4,6,...)` is MAX ignoring errors, which matters because half these columns are deliberately `NA()`.

**A scaled copy of every drawn column**, which is what the chart actually reads:

```
=C11/$O$11
```

Read from `C03A!Q11`.

**Axis bounds divided by the same span.** Dividing both the values and the bounds by one number is visually a no-op at today's figures and follows the data at any others. Zero divided by anything is zero, so the zero line, the reference rules and every caption positioned from plot geometry stay exactly where they were.

Two consequences worth knowing before you copy this:

- **The scaled columns are hidden, not parked far right.** A hidden column measures zero width, so the machinery can be added to a sheet without moving a single chart on it.
- **It cannot be used where an axis is read.** Every chart in this library hides its value axis except the XY pair, which show gridlines and tick labels - and dividing those by a span would print `0.25` where the data says `29.16%`. Those two are fitted at build time instead: correct for whatever data is present when the sheet is built, but not live on edit. That is a stated limit, not an oversight.

### 5. Text columns, because a linked label ignores its own format

A data label linked to a cell shows the **cell's** value and ignores the label's number format, so a variance of `31.61057692` prints every one of those digits. The formatting therefore has to happen in a cell:

```
=IF(ISNA(D11),"",TEXT(D11,"0"))
```

Read from `C03A!S11`.

The label is then linked to that column. The cell says exactly what the chart says, which is also why the two can be checked against each other.

### 6. Variances, and the relative variance that has no answer

```
=D11-C11
```

Read from `C03A!E11`.

A relative variance divides by the reference, which may be zero, and `NA()` is the honest answer - a bar of nothing, not a bar of zero:

```
=IF(C11=0,NA(),(D11-C11)/C11*100)
```

Read from `C03A!F11`.

### 7. The waterfall cascade

A waterfall is not a chart type here; it is two columns of arithmetic and an invisible series. Each row carries a **running level** and the level it **starts from**, and the sign column says whether the row adds or subtracts:

```
=G11+B12*C12
```

Read from `C12A!G12`. The running level: the previous level plus this row's signed value.

The floating bar is then the difference between the two, drawn over an invisible series holding the start. A subtotal row breaks the chain by reading the level directly instead of adding to it - which is what makes a subtotal a column from zero rather than a floating step.

### 8. A total is data

A total with components is a formula, never an input. A total that does not follow its parts is the same lie a static variance colour is:

```
=SUM(B10:B14)
```

Read from `C01A!B15`.

### 9. Scenario fills, the forecast hatch, and pins

These are formatting rather than formulas, so they are not readable out of the cells - but they are the notation, so they belong here:

| What | How it is drawn |
|---|---|
| AC - actual | solid dark fill |
| PY - prior year | solid light fill |
| PL / BU - plan or budget | outlined, no fill |
| FC - forecast | hatched, 45 degrees ascending |
| absolute variance | a column or bar, red or green **by impact** |
| relative variance | a **pin**: a thin stem with a head marker |

A pin has no chart type. It is a line series with the line hidden, markers on, and a **custom Y error bar** as the stem - the error bar's weight is in points, so the stem is exactly as thin as it should be, where a very narrow column bottoms out around 9px because `GapWidth` caps at 500. The head carries the minuend's scenario fill, which is what says whether a relative variance was measured or forecast.

Where a variance runs off the end of its panel, the bar is **clipped** and the label carries the overflow, because a bar drawn at full length off the plot says nothing about having been cut:

```
=IF(ISNA(F11),NA(),MEDIAN(-31.5,F11,238.5))
```

Read from `C12A!AP11`.

### 10. Page setup

Every sheet prints on one page: print area set to the chart zone alone, `FitToPagesWide/Tall = 1`, worksheet gridlines off. The data zone is deliberately outside the print area - it is the input, not the deliverable.

### The traps that fail silently

Each of these produces a wrong-looking sheet rather than an error.

| Trap | What you see |
|---|---|
| `""` where `NA()` was meant | a bar of nothing sitting on the axis |
| `OR()` guarding a lookup | `#REF!` across every spacer row - `OR` evaluates all its arguments, nested `IF`s short-circuit |
| a data label typed rather than linked | it stops tracking the cell the first time a figure changes |
| a label linked to a raw cell | full floating-point precision printed beside a rounded chart |
| `AutoFit` on the data block | charts drawn on top of the source data |
| an axis bound typed as a number | the chart clips the day somebody pastes their own figures |
| variance colour taken from the sign | every cost overrun green |

---

## Sheet by sheet

| Sheet | Family | Formulas | Typed | Charts |
|---|---|---|---|---|
| [C01A](#c01a) | structure | 54 | 20 | 1 |
| [C02A](#c02a) | structure | 173 | 57 | 1 |
| [C03A](#c03a) | tier stack | 279 | 24 | 1 |
| [C04A](#c04a) | tier stack | 364 | 38 | 1 |
| [C05X](#c05x) | tier stack | 688 | 25 | 1 |
| [C06F](#c06f) | tier stack | 643 | 33 | 2 |
| [C12A](#c12a) | tier stack | 859 | 48 | 1 |

### C01A

> Planned net sales in 2026 will increase by 12.2 kEUR (+12%) mainly due to software growth of 9.1 kEUR (+23%)

**Structure.** Categories stacked inside a column or bar, on one scale shared by every panel on the sheet.

| | |
|---|---|
| Cells used | `A1:E53` |
| Print area | `'C01A'!$G$1:$L$71` |
| Formula cells | 54, in 5 shapes |
| Typed number cells | 20 |

**Typed values** - `B10:E14`. Everything else on the sheet is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in this range, so go by the shading rather than by the count.

**The formulas behind everything else.** Cells that say the same thing about their own position are one row here; the example is the first of them, in ordinary A1.

| Cells | Named | Example | Formula |
|---|---|---|---|
| `B29:E34` | - | `B29` | `=IF(ISBLANK(B10),NA(),B10/$B$17)` |
| `B48:E53` | - | `B48` | `=IF(ISBLANK(B10),"",IF(B10<3.6,"",TEXT(B10,"0.0")))` |
| `B15:E15` | Total | `B15` | `=SUM(B10:B14)` |
| `B7` | Subject | `B7` | `=B2&IF(B3="",""," in "&B3)` |
| `B17` | span | `B17` | `=MAX(AGGREGATE(4,6,B15:E15),0.000000001)` |

**Chart objects**, in points from the top left of the sheet. The positions are what make separate charts read as one figure, so they are declared rather than dragged.

| Name | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|
| `panel_area` | 407.5 | 81 | 240 | 430 | 6 | 0 .. 1.1776 |

An axis range that is not round - `0.9` rather than `900` - is the scale block at work: the bounds were divided by the span cell, exactly as the values were. Multiply an axis bound by its span and you get back the range the geometry was measured in.

### C02A

> In Europe we achieved 3 098 kCHF (83%) of worldwide net sales (3 733 kCHF), USA net sales of 287 kCHF presents the biggest share outside of Europe (8%)

**Structure.** Categories stacked inside a column or bar, on one scale shared by every panel on the sheet.

| | |
|---|---|
| Cells used | `A1:T47` |
| Print area | `'C02A'!$W$1:$AG$73` |
| Formula cells | 173, in 5 shapes |
| Typed number cells | 57 |

**Typed values** - `B10:T12`. Everything else on the sheet is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in this range, so go by the shading rather than by the count.

**The formulas behind everything else.** Cells that say the same thing about their own position are one row here; the example is the first of them, in ordinary A1.

| Cells | Named | Example | Formula |
|---|---|---|---|
| `B27:T30` | - | `B27` | `=IF(ISBLANK(B10),NA(),B10/$B$15)` |
| `B44:T47` | - | `B44` | `=IF(ISBLANK(B10),"",IF(B10<15,"",TEXT(B10,"# ##0")))` |
| `B13:T13` | Total | `B13` | `=SUM(B10:B12)` |
| `B7` | Subject | `B7` | `=B2&IF(B3="",""," in "&B3)` |
| `B15` | span | `B15` | `=MAX(AGGREGATE(4,6,$B$13:$T$13),0.000000001)` |

**Chart objects**, in points from the top left of the sheet. The positions are what make separate charts read as one figure, so they are declared rather than dragged.

| Name | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|
| `panel_channel` | 1175.5 | 81 | 506.7 | 520 | 4 | 0 .. 1.0081 |

An axis range that is not round - `0.9` rather than `900` - is the scale block at work: the bounds were divided by the span cell, exactly as the values were. Multiply an axis bound by its span and you get back the range the geometry was measured in.

### C03A

> During the next two months the contribution will be below previous year but we expect an annual growth of 81 kEUR (+4.2%) for the full year

**Tier stack.** One chart per tier, stacked along the page and sharing a category axis. The tiers are separate chart objects positioned to line up, because one chart cannot carry two value axes that both start at zero.

| | |
|---|---|
| Cells used | `A1:AC22` |
| Print area | `'C03A'!$AD$1:$AO$21` |
| Hidden engine columns | `O` |
| Formula cells | 279, in 23 shapes |
| Typed number cells | 24 |

**Typed values** - `C11:D22`. Everything else on the sheet is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in this range, so go by the shading rather than by the count.

**The formulas behind everything else.** Cells that say the same thing about their own position are one row here; the example is the first of them, in ordinary A1.

| Cells | Named | Example | Formula |
|---|---|---|---|
| `Q11:R22` | ref scaled, measure scaled | `Q11` | `=C11/$O$11` |
| `V11:W22` | abs_up scaled, abs_dn scaled | `V11` | `=I11/$O$11` |
| `AB11:AC22` | rel_up_len scaled, rel_dn_len scaled | `AB11` | `=M11/$P$11` |
| `E11:E22` | ΔPY | `E11` | `=D11-C11` |
| `F11:F22` | ΔPY% | `F11` | `=IF(C11=0,NA(),(D11-C11)/C11*100)` |
| `G11:G22` | Measure AC | `G11` | `=IF($B11="AC",$D11,NA())` |
| `H11:H22` | Measure FC | `H11` | `=IF($B11="AC",NA(),$D11)` |
| `I11:I22` | ΔPY up | `I11` | `=IF($E11>0,$E11,NA())` |
| `J11:J22` | ΔPY down | `J11` | `=IF($E11<0,$E11,NA())` |
| `K11:K22` | ΔPY% up | `K11` | `=IF($F11>0,$F11,NA())` |
| `L11:L22` | ΔPY% down | `L11` | `=IF($F11<0,$F11,NA())` |
| `M11:M22` | up length | `M11` | `=IF($F11>0,$F11,0)` |
| `N11:N22` | down length | `N11` | `=IF($F11<0,-$F11,0)` |
| `S11:S22` | measure text | `S11` | `=IF(ISNA(D11),"",TEXT(D11,"0"))` |
| `T11:T22` | var_abs scaled | `T11` | `=E11/$O$11` |
| `U11:U22` | var_abs text | `U11` | `=IF(ISNA(E11),"",TEXT(E11,"+0;-0"))` |
| `X11:X22` | rel_up scaled | `X11` | `=K11/$P$11` |
| `Y11:Y22` | rel_up text | `Y11` | `=IF(ISNA(K11),"",TEXT(K11,"+0.0;-0.0"))` |
| `Z11:Z22` | rel_dn scaled | `Z11` | `=L11/$P$11` |
| `AA11:AA22` | rel_dn text | `AA11` | `=IF(ISNA(L11),"",TEXT(L11,"+0.0;-0.0"))` |
| `B7` | Subject | `B7` | `=B2&" in "&B3` |
| `O11` | span unit | `O11` | `=MAX(AGGREGATE(4,6,C11:C22,D11:D22,E11:E22,I11:I22,J11:J22),-AGGREGATE(5,6,C11:C22,D11:D22,E11:E22,I11:I22,J11:J22),0.000000001)` |
| `P11` | span rel | `P11` | `=MAX(AGGREGATE(4,6,K11:K22,L11:L22,M11:M22,N11:N22),-AGGREGATE(5,6,K11:K22,L11:L22,M11:M22,N11:N22),0.000000001)` |

**Chart objects**, in points from the top left of the sheet. The positions are what make separate charts read as one figure, so they are declared rather than dragged.

| Name | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|
| `tier_measure` | 683 | 86 | 520 | 244 | 2 | 0 .. 1.09 |

An axis range that is not round - `0.9` rather than `900` - is the scale block at work: the bounds were divided by the span cell, exactly as the values were. Multiply an axis bound by its span and you get back the range the geometry was measured in.

### C04A

> Compared to plan California (+34 kUSD) and Ohio (+31 kUSD) have the greatest absolute positive variances in net sales

**Tier stack.** One chart per tier, stacked along the page and sharing a category axis. The tiers are separate chart objects positioned to line up, because one chart cannot carry two value axes that both start at zero.

| | |
|---|---|
| Cells used | `A1:Y29` |
| Print area | `'C04A'!$Z$1:$AI$32` |
| Hidden engine columns | `M` |
| Formula cells | 364, in 19 shapes |
| Typed number cells | 38 |

**Typed values** - `C11:D29`. Everything else on the sheet is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in this range, so go by the shading rather than by the count.

**The formulas behind everything else.** Cells that say the same thing about their own position are one row here; the example is the first of them, in ordinary A1.

| Cells | Named | Example | Formula |
|---|---|---|---|
| 57 cells in `P11:U29` | measure scaled, abs_up scaled, abs_dn scaled | `P11` | `=C11/$M$11` |
| `X11:Y29` | rel_up scaled, rel_dn scaled | `X11` | `=I11/$N$11` |
| `E11:E29` | PL | `E11` | `=C11-D11` |
| `F11:F29` | ΔPL% | `F11` | `=IF(E11=0,NA(),D11/E11*100)` |
| `G11:G29` | ΔPL up | `G11` | `=IF($D11>0,$D11,NA())` |
| `H11:H29` | ΔPL down | `H11` | `=IF($D11<0,$D11,NA())` |
| `I11:I29` | ΔPL% up | `I11` | `=IF($F11>0,$F11,NA())` |
| `J11:J29` | ΔPL% down | `J11` | `=IF($F11<0,$F11,NA())` |
| `K11:K29` | up length | `K11` | `=IF($F11>0,$F11,0)` |
| `L11:L29` | down length | `L11` | `=IF($F11<0,-$F11,0)` |
| `O11:O29` | ref scaled | `O11` | `=E11/$M$11` |
| `Q11:Q29` | measure text | `Q11` | `=IF(ISNA(C11),"",TEXT(C11,"0"))` |
| `R11:R29` | var_abs scaled | `R11` | `=D11/$M$11` |
| `S11:S29` | var_abs text | `S11` | `=IF(ISNA(D11),"",TEXT(D11,"+0;-0"))` |
| `V11:V29` | var_rel scaled | `V11` | `=F11/$N$11` |
| `W11:W29` | var_rel text | `W11` | `=IF(ISNA(F11),"",TEXT(F11,"+0;-0"))` |
| `B7` | Subject | `B7` | `=B2&" in "&B3` |
| `M11` | span unit | `M11` | `=MAX(AGGREGATE(4,6,E11:E29,C11:C29,D11:D29,G11:G29,H11:H29),-AGGREGATE(5,6,E11:E29,C11:C29,D11:D29,G11:G29,H11:H29),0.000000001)` |
| `N11` | span rel | `N11` | `=MAX(AGGREGATE(4,6,F11:F29,I11:I29,J11:J29),-AGGREGATE(5,6,F11:F29,I11:I29,J11:J29),0.000000001)` |

**Chart objects**, in points from the top left of the sheet. The positions are what make separate charts read as one figure, so they are declared rather than dragged.

| Name | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|
| `tier_measure` | 631.5 | 86 | 430 | 400 | 2 | 0 .. 240 |

An axis range that is not round - `0.9` rather than `900` - is the scale block at work: the bounds were divided by the span cell, exactly as the values were. Multiply an axis bound by its span and you get back the range the geometry was measured in.

### C05X

> We expect a plus of 24 kEUR (+15.6%) vs plan until end of the year because of the positive forecast beginning in September

**Tier stack.** One chart per tier, stacked along the page and sharing a category axis. The tiers are separate chart objects positioned to line up, because one chart cannot carry two value axes that both start at zero.

| | |
|---|---|
| Cells used | `A1:AU27` |
| Print area | `'C05X'!$AV$1:$BI$52` |
| Hidden engine columns | `W` |
| Formula cells | 688, in 60 shapes |
| Typed number cells | 25 |

**Typed values** - 25 cells in `D11:E24`. Everything else on the sheet is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in this range, so go by the shading rather than by the count.

**The formulas behind everything else.** Cells that say the same thing about their own position are one row here; the example is the first of them, in ordinary A1.

| Cells | Named | Example | Formula |
|---|---|---|---|
| 51 cells in `AC11:AH27` | bar_ac scaled, wf_base scaled, wf_up_ac scaled | `AC11` | `=J11/$W$11` |
| 40 cells in `G11:T27` | ΔPL%, PY bar, PL bar, AC bar | `G11` | `=NA()` |
| 34 cells in `AE11:AJ27` | bar_fc scaled, wf_dn_ac scaled | `AE11` | `=K11/$W$11` |
| `AT11:AU27` | rel_up_len scaled, rel_dn_len scaled | `AT11` | `=U11/$X$11` |
| 33 cells in `F11:V27` | PL, level, base, up AC | `F11` | `=0` |
| `Y11:Y27` | bar_py scaled | `Y11` | `=H11/$W$11` |
| `Z11:Z27` | bar_py text | `Z11` | `=IF(ISNA(H11),"",TEXT(H11,"#,##0"))` |
| `AA11:AA27` | bar_pl scaled | `AA11` | `=I11/$W$11` |
| `AB11:AB27` | bar_pl text | `AB11` | `=IF(ISNA(I11),"",TEXT(I11,"#,##0"))` |
| `AD11:AD27` | bar_ac text | `AD11` | `=IF(ISNA(J11),"",TEXT(J11,"#,##0"))` |
| `AF11:AF27` | bar_fc text | `AF11` | `=IF(ISNA(K11),"",TEXT(K11,"#,##0"))` |
| `AI11:AI27` | wf_up_ac text | `AI11` | `=IF(ISNA(O11),"",TEXT(O11,"+#,##0"))` |
| `AK11:AK27` | wf_dn_ac text | `AK11` | `=IF(ISNA(P11),"",TEXT(P11,"""-""#,##0"))` |
| `AL11:AL27` | wf_up_fc scaled | `AL11` | `=Q11/$W$11` |
| `AM11:AM27` | wf_up_fc text | `AM11` | `=IF(ISNA(Q11),"",TEXT(Q11,"+#,##0"))` |
| `AN11:AN27` | wf_dn_fc scaled | `AN11` | `=R11/$W$11` |
| `AO11:AO27` | wf_dn_fc text | `AO11` | `=IF(ISNA(R11),"",TEXT(R11,"""-""#,##0"))` |
| `AP11:AP27` | rel_up scaled | `AP11` | `=S11/$X$11` |
| `AQ11:AQ27` | rel_up text | `AQ11` | `=IF(ISNA(S11),"",TEXT(S11,"+0.0;-0.0"))` |
| `AR11:AR27` | rel_dn scaled | `AR11` | `=T11/$X$11` |
| `AS11:AS27` | rel_dn text | `AS11` | `=IF(ISNA(T11),"",TEXT(T11,"+0.0;-0.0"))` |
| `S13:S25` | Jan | `S13` | `=IF(G13>0,G13,NA())` |
| `T13:T25` | Jan | `T13` | `=IF(G13<0,G13,NA())` |
| `U13:U25` | Jan | `U13` | `=IF(G13>0,G13,0)` |
| `V13:V25` | Jan | `V13` | `=IF(G13<0,-G13,0)` |
| `F13:F24` | Jan | `F13` | `=D13-E13` |
| `G13:G24` | Jan | `G13` | `=IF(F13=0,NA(),E13/F13*100)` |
| `I13:I24` | Jan | `I13` | `=F13` |
| `J13:J24` | Jan | `J13` | `=IF($B13="AC",D13,NA())` |
| `K13:K24` | Jan | `K13` | `=IF($B13="FC",D13,NA())` |
| `L13:L24` | Jan | `L13` | `=M13+E13` |
| `M13:M24` | Jan | `M13` | `=L12` |
| `N13:N24` | Jan | `N13` | `=MIN(M13,L13)` |
| `O13:O24` | Jan | `O13` | `=IF(AND($B13="AC",E13>0),E13,0)` |
| `P13:P24` | Jan | `P13` | `=IF(AND($B13="AC",E13<0),-E13,0)` |
| `Q13:Q24` | Jan | `Q13` | `=IF(AND($B13="FC",E13>0),E13,0)` |
| `R13:R24` | Jan | `R13` | `=IF(AND($B13="FC",E13<0),-E13,0)` |
| 3 cells in `M11:M25` | from | `M11` | `=L11` |
| `L26:L27` | vs PL | `L26` | `=$D$25` |
| `O26:O27` | vs PL | `O26` | `=IF(E26>0,E26,0)` |

*20 further named shapes are not listed - this sheet has 60 in all.*

**Chart objects**, in points from the top left of the sheet. The positions are what make separate charts read as one figure, so they are declared rather than dragged.

| Name | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|
| `tier_wf` | 1038.5 | 86 | 620 | 240 | 5 | 140 .. 200 |

An axis range that is not round - `0.9` rather than `900` - is the scale block at work: the bounds were divided by the span cell, exactly as the values were. Multiply an axis bound by its span and you get back the range the geometry was measured in.

### C06F

> In Q3 2025, the total decrease in net sales compared to PY was 343 kUSD

**Tier stack.** One chart per tier, stacked along the page and sharing a category axis. The tiers are separate chart objects positioned to line up, because one chart cannot carry two value axes that both start at zero.

| | |
|---|---|
| Cells used | `A1:AM30` |
| Print area | `'C06F'!$AN$1:$BC$35` |
| Hidden engine columns | `T` |
| Formula cells | 643, in 45 shapes |
| Typed number cells | 33 |

**Typed values** - 33 cells in `D11:E29`. Everything else on the sheet is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in this range, so go by the shading rather than by the count.

**The formulas behind everything else.** Cells that say the same thing about their own position are one row here; the example is the first of them, in ordinary A1.

| Cells | Named | Example | Formula |
|---|---|---|---|
| 82 cells in `F11:Q30` | PY, AC bar, PY short, PL bar | `F11` | `=0` |
| 60 cells in `AA11:AF30` | bar_py scaled, wf_base scaled, wf_up scaled | `AA11` | `=K11/$T$11` |
| `X11:Y30` | bar_short scaled, bar_pl scaled | `X11` | `=I11/$T$11` |
| 40 cells in `AC11:AH30` | bar_ac scaled, wf_dn scaled | `AC11` | `=L11/$T$11` |
| `AL11:AM30` | rel_up scaled, rel_dn scaled | `AL11` | `=R11/$U$11` |
| `V11:V30` | bar_state scaled | `V11` | `=H11/$T$11` |
| `W11:W30` | bar_state text | `W11` | `=IF(ISNA(H11),"",TEXT(H11,"#,##0"))` |
| `Z11:Z30` | bar_pl text | `Z11` | `=IF(ISNA(J11),"",TEXT(J11,"#,##0"))` |
| `AB11:AB30` | bar_py text | `AB11` | `=IF(ISNA(K11),"",TEXT(K11,"#,##0"))` |
| `AD11:AD30` | bar_ac text | `AD11` | `=IF(ISNA(L11),"",TEXT(L11,"#,##0"))` |
| `AG11:AG30` | wf_up text | `AG11` | `=IF(ISNA(P11),"",TEXT(P11,"+#,##0"))` |
| `AI11:AI30` | wf_dn text | `AI11` | `=IF(ISNA(Q11),"",TEXT(Q11,"""-""#,##0"))` |
| `AJ11:AJ30` | var_rel scaled | `AJ11` | `=E11/$U$11` |
| `AK11:AK30` | var_rel text | `AK11` | `=IF(ISNA(E11),"",TEXT(E11,"+0;-0"))` |
| 17 cells in `P13:P30` | Iowa | `P13` | `=IF(G13>0,G13,0)` |
| 17 cells in `Q13:Q30` | Iowa | `Q13` | `=IF(G13<0,-G13,0)` |
| `R13:R28` | Iowa | `R13` | `=IF(E13>0,E13,NA())` |
| `S13:S28` | Iowa | `S13` | `=IF(E13<0,E13,NA())` |
| `F13:F27` | Iowa | `F13` | `=IF(E13=-100,NA(),D13/(1+E13/100))` |
| `G13:G27` | Iowa | `G13` | `=D13-F13` |
| `H13:H27` | Iowa | `H13` | `=D13` |
| `I13:I27` | Iowa | `I13` | `=MAX(F13-D13,0)` |
| `M13:M27` | Iowa | `M13` | `=N13+G13` |
| `N13:N27` | Iowa | `N13` | `=M12` |
| `O13:O27` | Iowa | `O13` | `=MIN(N13,M13)` |
| 8 cells in `R11:S30` | ΔPY% up, ΔPY% down | `R11` | `=NA()` |
| 3 cells in `N11:N28` | from | `N11` | `=M11` |
| 2 cells in `M29:N30` | ΔPY | `N29` | `=$D$28` |
| `B7` | Subject | `B7` | `=B2&" in "&B3` |
| `J11` | PL bar | `J11` | `=D11` |
| `T11` | span unit | `T11` | `=MAX(AGGREGATE(4,6,H11:H30,I11:I30,J11:J30,K11:K30,L11:L30,O11:O30,P11:P30,Q11:Q30),-AGGREGATE(5,6,H11:H30,I11:I30,J11:J30,K11:K30,L11:L30,O11:O30,P11:P30,Q11:Q30),0.000000001)` |
| `U11` | span rel | `U11` | `=MAX(AGGREGATE(4,6,E11:E30,R11:R30,S11:S30),-AGGREGATE(5,6,E11:E30,R11:R30,S11:S30),0.000000001)` |
| `D12` | PY | `D12` | `=SUM(F13:F27)` |
| `K12` | PY | `K12` | `=D12` |
| `M12` | PY | `M12` | `=D12` |
| `D28` | AC | `D28` | `=SUM(D13:D27)` |
| `G28` | AC | `G28` | `=D28-$D$12` |
| `L28` | AC | `L28` | `=D28` |
| `M28` | AC | `M28` | `=M27` |
| `G29` | ΔPY | `G29` | `=$D$28-$D$12` |

*5 further named shapes are not listed - this sheet has 45 in all.*

**Chart objects**, in points from the top left of the sheet. The positions are what make separate charts read as one figure, so they are declared rather than dragged.

| Name | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|
| `tier_measure` | 970.5 | 86 | 560 | 430 | 5 | 0 .. 2200 |
| `tier_wf` | 1540.5 | 86 | 140 | 430 | 3 | 1600 .. 2200 |

An axis range that is not round - `0.9` rather than `900` - is the scale block at work: the bounds were divided by the span cell, exactly as the values were. Multiply an axis bound by its span and you get back the range the geometry was measured in.

### C12A

> Compared to 2024, the higher operating expenses (+187 kEUR) were mainly compensated by higher license sales (+183 kEUR), leading to a higher group result (+91 kEUR)

**Tier stack.** One chart per tier, stacked along the page and sharing a category axis. The tiers are separate chart objects positioned to line up, because one chart cannot carry two value axes that both start at zero.

| | |
|---|---|
| Cells used | `A1:AU30` |
| Print area | `'C12A'!$AV$1:$BD$33` |
| Hidden engine columns | `AP`, `AU`, `Y` |
| Formula cells | 859, in 44 shapes |
| Typed number cells | 48 |

**Typed values** - 48 cells in `B11:D30`. Everything else on the sheet is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in this range, so go by the shading rather than by the count.

**The formulas behind everything else.** Cells that say the same thing about their own position are one row here; the example is the first of them, in ordinary A1.

| Cells | Named | Example | Formula |
|---|---|---|---|
| 80 cells in `Z11:AF30` | py_base scaled, py_add scaled, ac_base scaled, ac_add scaled | `Z11` | `=I11/$Y$11` |
| `AL11:AO30` | abs_pos_good scaled, abs_pos_bad scaled, abs_neg_good scaled, abs_neg_bad scaled | `AL11` | `=Q11/$Y$11` |
| `AQ11:AT30` | rel_pos_good clipped, rel_pos_bad clipped, rel_neg_good clipped, rel_neg_bad clipped | `AQ11` | `=IF(ISNA(U11),NA(),MEDIAN(-31.5,U11,238.5))` |
| 40 cells in `I11:N30` | PY base, AC base | `I11` | `=MIN(H11,G11)` |
| 40 cells in `AB11:AG30` | py_add text, ac_add text | `AB11` | `=IF(ISNA(J11),"",TEXT(J11,"#,##0"))` |
| 40 cells in `AC11:AH30` | py_sub scaled, ac_sub scaled | `AC11` | `=K11/$Y$11` |
| 40 cells in `AD11:AI30` | py_sub text, ac_sub text | `AD11` | `=IF(ISNA(K11),"",TEXT(K11,"#,##0"))` |
| 26 cells in `H12:M29` | + Consulting | `H12` | `=G11` |
| `E11:E30` | ΔPY | `E11` | `=D11-C11` |
| `F11:F30` | ΔPY% | `F11` | `=IF(C11=0,NA(),(D11-C11)/C11*100)` |
| `J11:J30` | PY adds | `J11` | `=IF(B11>0,ABS(G11-H11),0)` |
| `K11:K30` | PY subs | `K11` | `=IF(B11<0,ABS(G11-H11),0)` |
| `O11:O30` | AC adds | `O11` | `=IF(B11>0,ABS(L11-M11),0)` |
| `P11:P30` | AC subs | `P11` | `=IF(B11<0,ABS(L11-M11),0)` |
| `Q11:Q30` | ΔPY up ok | `Q11` | `=IF(AND($E11>0,$B11>0),$E11,NA())` |
| `R11:R30` | ΔPY up bad | `R11` | `=IF(AND($E11>0,$B11<0),$E11,NA())` |
| `S11:S30` | ΔPY dn ok | `S11` | `=IF(AND($E11<0,$B11<0),$E11,NA())` |
| `T11:T30` | ΔPY dn bad | `T11` | `=IF(AND($E11<0,$B11>0),$E11,NA())` |
| `U11:U30` | ΔPY% up ok | `U11` | `=IF(AND($F11>0,$B11>0),$F11,NA())` |
| `V11:V30` | ΔPY% up bad | `V11` | `=IF(AND($F11>0,$B11<0),$F11,NA())` |
| `W11:W30` | ΔPY% dn ok | `W11` | `=IF(AND($F11<0,$B11<0),$F11,NA())` |
| `X11:X30` | ΔPY% dn bad | `X11` | `=IF(AND($F11<0,$B11>0),$F11,NA())` |
| `AJ11:AJ30` | var_abs scaled | `AJ11` | `=E11/$Y$11` |
| `AK11:AK30` | var_abs text | `AK11` | `=IF(ISNA(E11),"",TEXT(E11,"+0;-0"))` |
| `AP11:AP30` | var_rel clipped | `AP11` | `=IF(ISNA(F11),NA(),MEDIAN(-31.5,F11,238.5))` |
| `AU11:AU30` | ΔPY% text | `AU11` | `=IF(ISNA(F11),"",IF(F11>265,TEXT(F11,"+0;-0")&REPT(UNICHAR(9658),MIN(3,MAX(1,INT(F11/265)))),IF(F11<-35,REPT(UNICHAR(9668),MIN(3,MAX(1,INT(-F11/35))))&TEXT(F11,"+0;-0"),TEXT(F11,"+0;-0"))))` |
| 13 cells in `G12:G29` | + Consulting | `G12` | `=G11+B12*C12` |
| 13 cells in `L12:L29` | + Consulting | `L12` | `=L11+B12*D12` |
| 12 cells in `H11:M30` | PY from, AC from | `H11` | `=0` |
| 12 cells in `G15:L30` | - | `G15` | `=G14` |
| 5 cells in `C15:C30` | - | `C15` | `=G14` |
| 5 cells in `D15:D30` | - | `D15` | `=L14` |
| 2 cells in `H22:M22` | - Operating expenses | `H22` | `=G16` |
| `B7` | Subject | `B7` | `=B2&" in "&B3` |
| `G11` | PY level | `G11` | `=0+B11*C11` |
| `L11` | AC level | `L11` | `=0+B11*D11` |
| `Y11` | span unit | `Y11` | `=MAX(AGGREGATE(4,6,I11:I30,J11:J30,K11:K30,N11:N30,O11:O30,P11:P30,E11:E30,Q11:Q30,R11:R30,S11:S30,T11:T30),-AGGREGATE(5,6,I11:I30,J11:J30,K11:K30,N11:N30,O11:O30,P11:P30,E11:E30,Q11:Q30,R11:R30,S11:S30,T11:T30),0.000000001)` |
| `A15` | + Other revenue | `A15` | `= Sales revenue` |
| `C22` | - Operating expenses | `C22` | `=ABS(G21-G16)` |
| `D22` | - Operating expenses | `D22` | `=ABS(L21-L16)` |

*4 further named shapes are not listed - this sheet has 44 in all.*

**Chart objects**, in points from the top left of the sheet. The positions are what make separate charts read as one figure, so they are declared rather than dragged.

| Name | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|
| `tier_wf_ac` | 1177.5 | 86 | 400 | 402 | 3 | 0 .. 1120 |

An axis range that is not round - `0.9` rather than `900` - is the scale block at work: the bounds were divided by the span cell, exactly as the values were. Multiply an axis bound by its span and you get back the range the geometry was measured in.

---

*Generated by `ibcs_doc.py` from the workbook itself.*
