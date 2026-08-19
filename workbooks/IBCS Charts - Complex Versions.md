# IBCS Charts - Complex Versions - how it is built

Every formula quoted here was **read back out of `IBCS Charts - Complex Versions.xlsx`** after it was written, so this document and that workbook cannot disagree. Where it says a formula sits in `C03A!E13`, it does.

| | |
|---|---|
| Sheets | 18 (a Read me and 17 templates) |
| Formula cells | 7119 |
| Typed cells | 1434 |
| Chart objects | 40 |
| Distinct formula shapes | 389 |

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

```
="Δ"&$B$5&"%"
```

Read from `C03A!B9`.

On a chart sheet the caption *cell* is then linked to a text box on the chart - two hops, one source. **Nothing a chart says is typed into the chart.** Type it in a cell and link the box to the cell, every time.

### 2. Two zones, and nothing straddles them

Each sheet is a data zone of declared width on the left and the charts to its right, with the chart zone's left edge measured from where the data ends. The print area is the chart zone alone:

| Sheet | Cells used | Print area |
|---|---|---|
| C01A | `A1:F104` | `'C01A'!$G$1:$R$36` |
| C02A | `A1:V47` | `'C02A'!$W$1:$AI$73` |
| C03A | `A1:AC24` | `'C03A'!$AD$1:$AO$37` |
| C04A | `A1:Y31` | `'C04A'!$Z$1:$AR$32` |
| C05X | `A1:AU28` | `'C05X'!$AV$1:$BI$46` |
| C06F | `A1:AM31` | `'C06F'!$AN$1:$BG$35` |
| C07C | `A1:M59` | `'C07C'!$N$1:$AA$36` |
| C08H | `A1:U65` | `'C08H'!$V$1:$AL$81` |
| C09C | `A1:E226` | `'C09C'!$F$1:$T$46` |
| C10D | `A1:E32` | `'C10D'!$F$1:$T$40` |
| C11A | `A1:H62` | `'C11A'!$I$4:$Z$88` |
| C12A | `A1:AU32` | `'C12A'!$AV$1:$BS$33` |
| C13D | `A1:BU106` | `'C13D'!$A$109:$N$148` |
| T01B | `A1:O35` | `'T01B'!$A$9:$O$35` |
| T02A | `A1:AB34` | `'T02A'!$A$9:$J$35` |
| T03A | `A1:H35` | `'T03A'!$A$9:$H$35` |
| T04A | `A1:P35` | `'T04A'!$A$9:$F$36` |

**A table sheet is the other way round.** On T01B, T02A, T03A, T04A the grid *is* the deliverable, so the print area is the grid - starting below the six typed title rows, which are excluded for exactly the reason a chart sheet excludes its data zone: they are the input.

C13D parks the charts **below** the data rather than beside it, because the panel engine owns its own columns to the right. Same rule, read the other way: the print area holds the deliverable and nothing else, so it goes wherever the data does not.

**Never `AutoFit` a data block a chart is positioned against.** One long string moves the boundary and the charts end up drawn over the source data. Column widths here are declared, never computed from content.

### 3. Splitting a series by what it means, not by its sign

A variance is drawn in two colours, and Excel gives one colour to a series. So the split happens in the *cells*: one column per direction, each `NA()` where the other owns the point.

```
=IF($E13>0,$E13,NA())
```

Read from `C03A!I13`.

`NA()` and not `""` - an empty string plots as a **zero**, which puts a bar of nothing on the axis where there should be no bar at all. That one substitution is the most common way to get a chart that looks broken for no visible reason.

The same split does the work for impact: within one variance column of a statement the favourable values sit on **both** sides of the axis, because a cost line up is adverse and a cost line down is favourable. Split by impact, never by sign, or every cost overrun comes out green.

### 4. The scale block - why the charts follow your numbers

This is the one technique with no equivalent in a hand-drawn chart, and it exists because **Excel will not bind an axis bound to a formula**. An axis maximum is a number you type; it cannot be `=MAX(...)`. So a chart built for one set of figures clips or shrinks the moment somebody pastes their own over the top.

The way round it is to invert the problem: instead of scaling the axis to the data, **scale the data to a fixed axis**. Three parts:

**A span cell** - one per group of tiers sharing a unit - holding the largest magnitude anywhere in that group:

```
=MAX(AGGREGATE(4,6,C13:C24,D13:D24,E13:E24,I13:I24,J13:J24),-AGGREGATE(5,6,C13:C24,D13:D24,E13:E24,I13:I24,J13:J24),0.000000001)
```

Read from `C03A!O13`. `AGGREGATE(4,6,...)` is MAX ignoring errors, which matters because half these columns are deliberately `NA()`.

**A scaled copy of every drawn column**, which is what the chart actually reads:

```
=C13/$O$13
```

Read from `C03A!Q13`.

**Axis bounds divided by the same span.** Dividing both the values and the bounds by one number is visually a no-op at today's figures and follows the data at any others. Zero divided by anything is zero, so the zero line, the reference rules and every caption positioned from plot geometry stay exactly where they were.

Two consequences worth knowing before you copy this:

- **The scaled columns are hidden, not parked far right.** A hidden column measures zero width, so the machinery can be added to a sheet without moving a single chart on it.
- **It cannot be used where an axis is read.** Every chart in this library hides its value axis except the XY pair, which show gridlines and tick labels - and dividing those by a span would print `0.25` where the data says `29.16%`. Those two are fitted at build time instead: correct for whatever data is present when the sheet is built, but not live on edit. That is a stated limit, not an oversight.

### 5. Text columns, because a linked label ignores its own format

A data label linked to a cell shows the **cell's** value and ignores the label's number format, so a variance of `31.61057692` prints every one of those digits. The formatting therefore has to happen in a cell:

```
=IF(ISNA(D13),"",TEXT(D13,"0"))
```

Read from `C03A!S13`.

The label is then linked to that column. The cell says exactly what the chart says, which is also why the two can be checked against each other.

### 6. Variances, and the relative variance that has no answer

```
=D13-C13
```

Read from `C03A!E13`.

A relative variance divides by the reference, which may be zero, and `NA()` is the honest answer - a bar of nothing, not a bar of zero:

```
=IF(C13=0,NA(),(D13-C13)/C13*100)
```

Read from `C03A!F13`.

### 7. The waterfall cascade

A waterfall is not a chart type here; it is two columns of arithmetic and an invisible series. Each row carries a **running level** and the level it **starts from**, and the sign column says whether the row adds or subtracts:

```
=G13+B14*C14
```

Read from `C12A!G14`. The running level: the previous level plus this row's signed value.

The floating bar is then the difference between the two, drawn over an invisible series holding the start. A subtotal row breaks the chain by reading the level directly instead of adding to it - which is what makes a subtotal a column from zero rather than a floating step.

### 8. A total is data

A total with components is a formula, never an input. A total that does not follow its parts is the same lie a static variance colour is:

```
=SUM(B15:B18)
```

Read from `T04A!B19`.

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
=IF(ISNA(F13),NA(),MEDIAN(-31.5,F13,238.5))
```

Read from `C12A!AP13`.

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
| [C01A](#c01a) | structure | 95 | 36 | 3 |
| [C02A](#c02a) | structure | 197 | 57 | 1 |
| [C03A](#c03a) | tier stack | 281 | 24 | 3 |
| [C04A](#c04a) | tier stack | 366 | 38 | 3 |
| [C05X](#c05x) | tier stack | 689 | 25 | 3 |
| [C06F](#c06f) | tier stack | 644 | 33 | 3 |
| [C07C](#c07c) | line | 183 | 36 | 1 |
| [C08H](#c08h) | line | 262 | 41 | 3 |
| [C09C](#c09c) | XY | 331 | 362 | 1 |
| [C10D](#c10d) | XY | 2 | 60 | 1 |
| [C11A](#c11a) | tree | 123 | 21 | 6 |
| [C12A](#c12a) | tier stack | 861 | 48 | 4 |
| [C13D](#c13d) | panel | 2124 | 415 | 2 |
| [T01B](#t01b) | table | 190 | 104 | 0 |
| [T02A](#t02a) | table | 424 | 64 | 4 |
| [T03A](#t03a) | table | 114 | 42 | 0 |
| [T04A](#t04a) | table | 233 | 28 | 2 |

### C01A

> Planned net sales in 2026 will increase by 12.2 kEUR (+12%) mainly due to software growth of 9.1 kEUR (+23%)

**Structure.** Categories stacked inside a column or bar, on one scale shared by every panel on the sheet.

| | |
|---|---|
| Cells used | `A1:F104` |
| Print area | `'C01A'!$G$1:$R$36` |
| Formula cells | 95, in 6 shapes |
| Typed number cells | 36 |

**Typed values** - 36 cells in `B10:F31`. Everything else on the sheet is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in this range, so go by the shading rather than by the count.

**The formulas behind everything else.** Cells that say the same thing about their own position are one row here; the example is the first of them, in ordinary A1.

| Cells | Named | Example | Formula |
|---|---|---|---|
| 43 cells in `B46:F68` | - | `B46` | `=IF(ISBLANK(B10),NA(),B10/$B$34)` |
| 43 cells in `B82:F104` | - | `B82` | `=IF(ISBLANK(B10),"",IF(B10<3.6,"",TEXT(B10,"0.0")))` |
| 6 cells in `B15:F32` | Total | `B15` | `=SUM(B10:B14)` |
| `B7` | Subject | `B7` | `=B2&IF(B3="",""," in "&B3)` |
| `B24` | Total | `B24` | `=SUM(B18:B23)` |
| `B34` | span | `B34` | `=MAX(AGGREGATE(4,6,B15:F15,B24:B24,B32:B32),0.000000001)` |

**Chart objects**, in points from the top left of the sheet. The positions are what make separate charts read as one figure, so they are declared rather than dragged.

| Name | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|
| `panel_area` | 407.5 | 81 | 300 | 430 | 6 | 0 .. 1.0517 |
| `panel_industry` | 731.5 | 81 | 96 | 430 | 7 | 0 .. 1.0517 |
| `panel_region` | 851.5 | 81 | 96 | 430 | 6 | 0 .. 1.0517 |

An axis range that is not round - `0.9` rather than `900` - is the scale block at work: the bounds were divided by the span cell, exactly as the values were. Multiply an axis bound by its span and you get back the range the geometry was measured in.

### C02A

> In Europe we achieved 3 098 kCHF (83%) of worldwide net sales (3 733 kCHF), USA net sales of 287 kCHF presents the biggest share outside of Europe (8%)

**Structure.** Categories stacked inside a column or bar, on one scale shared by every panel on the sheet.

| | |
|---|---|
| Cells used | `A1:V47` |
| Print area | `'C02A'!$W$1:$AI$73` |
| Formula cells | 197, in 7 shapes |
| Typed number cells | 57 |

**Typed values** - 57 cells in `B10:U12`. Everything else on the sheet is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in this range, so go by the shading rather than by the count.

**The formulas behind everything else.** Cells that say the same thing about their own position are one row here; the example is the first of them, in ordinary A1.

| Cells | Named | Example | Formula |
|---|---|---|---|
| `B27:V30` | - | `B27` | `=IF(ISBLANK(B10),NA(),B10/$B$15)` |
| `B44:V47` | - | `B44` | `=IF(ISBLANK(B10),"",IF(B10<15,"",TEXT(B10,"# ##0")))` |
| `B13:V13` | Total | `B13` | `=SUM(B10:B12)` |
| `P10:P12` | Europe | `P10` | `=SUM(B10:O10)` |
| `V10:V12` | World | `V10` | `=SUM(B10:O10)+SUM(Q10:U10)` |
| `B7` | Subject | `B7` | `=B2&IF(B3="",""," in "&B3)` |
| `B15` | span | `B15` | `=MAX(AGGREGATE(4,6,$B$13:$O$13,$Q$13:$U$13),0.000000001)` |

**Chart objects**, in points from the top left of the sheet. The positions are what make separate charts read as one figure, so they are declared rather than dragged.

| Name | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|
| `panel_channel` | 1175.5 | 81 | 560 | 520 | 4 | 0 .. 1.0081 |

An axis range that is not round - `0.9` rather than `900` - is the scale block at work: the bounds were divided by the span cell, exactly as the values were. Multiply an axis bound by its span and you get back the range the geometry was measured in.

### C03A

> During the next two months the contribution will be below previous year but we expect an annual growth of 81 kEUR (+4.2%) for the full year

**Tier stack.** One chart per tier, stacked along the page and sharing a category axis. The tiers are separate chart objects positioned to line up, because one chart cannot carry two value axes that both start at zero.

| | |
|---|---|
| Cells used | `A1:AC24` |
| Print area | `'C03A'!$AD$1:$AO$37` |
| Hidden engine columns | `O` |
| Formula cells | 281, in 25 shapes |
| Typed number cells | 24 |

**Typed values** - `C13:D24`. Everything else on the sheet is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in this range, so go by the shading rather than by the count.

**The formulas behind everything else.** Cells that say the same thing about their own position are one row here; the example is the first of them, in ordinary A1.

| Cells | Named | Example | Formula |
|---|---|---|---|
| `Q13:R24` | ref scaled, measure scaled | `Q13` | `=C13/$O$13` |
| `V13:W24` | abs_up scaled, abs_dn scaled | `V13` | `=I13/$O$13` |
| `AB13:AC24` | rel_up_len scaled, rel_dn_len scaled | `AB13` | `=M13/$P$13` |
| `E13:E24` | ΔPY | `E13` | `=D13-C13` |
| `F13:F24` | ΔPY% | `F13` | `=IF(C13=0,NA(),(D13-C13)/C13*100)` |
| `G13:G24` | Measure AC | `G13` | `=IF($B13="AC",$D13,NA())` |
| `H13:H24` | Measure FC | `H13` | `=IF($B13="AC",NA(),$D13)` |
| `I13:I24` | ΔPY up | `I13` | `=IF($E13>0,$E13,NA())` |
| `J13:J24` | ΔPY down | `J13` | `=IF($E13<0,$E13,NA())` |
| `K13:K24` | ΔPY% up | `K13` | `=IF($F13>0,$F13,NA())` |
| `L13:L24` | ΔPY% down | `L13` | `=IF($F13<0,$F13,NA())` |
| `M13:M24` | up length | `M13` | `=IF($F13>0,$F13,0)` |
| `N13:N24` | down length | `N13` | `=IF($F13<0,-$F13,0)` |
| `S13:S24` | measure text | `S13` | `=IF(ISNA(D13),"",TEXT(D13,"0"))` |
| `T13:T24` | var_abs scaled | `T13` | `=E13/$O$13` |
| `U13:U24` | var_abs text | `U13` | `=IF(ISNA(E13),"",TEXT(E13,"+0;-0"))` |
| `X13:X24` | rel_up scaled | `X13` | `=K13/$P$13` |
| `Y13:Y24` | rel_up text | `Y13` | `=IF(ISNA(K13),"",TEXT(K13,"+0.0;-0.0"))` |
| `Z13:Z24` | rel_dn scaled | `Z13` | `=L13/$P$13` |
| `AA13:AA24` | rel_dn text | `AA13` | `=IF(ISNA(L13),"",TEXT(L13,"+0.0;-0.0"))` |
| `B7` | Subject | `B7` | `=B2&" in "&B3` |
| `B9` | var_rel caption | `B9` | `="Δ"&$B$5&"%"` |
| `B10` | var_abs caption | `B10` | `="Δ"&$B$5&" "&$B$3` |
| `O13` | span unit | `O13` | `=MAX(AGGREGATE(4,6,C13:C24,D13:D24,E13:E24,I13:I24,J13:J24),-AGGREGATE(5,6,C13:C24,D13:D24,E13:E24,I13:I24,J13:J24),0.000000001)` |
| `P13` | span rel | `P13` | `=MAX(AGGREGATE(4,6,K13:K24,L13:L24,M13:M24,N13:N24),-AGGREGATE(5,6,K13:K24,L13:L24,M13:M24,N13:N24),0.000000001)` |

**Chart objects**, in points from the top left of the sheet. The positions are what make separate charts read as one figure, so they are declared rather than dragged.

| Name | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|
| `tier_var_rel` | 683 | 96 | 520 | 104 | 2 | -0.8671 .. 1.3295 |
| `tier_var_abs` | 683 | 204 | 520 | 104 | 3 | -0.173 .. 0.263 |
| `tier_measure` | 683 | 312 | 520 | 244 | 2 | 0 .. 1.09 |

An axis range that is not round - `0.9` rather than `900` - is the scale block at work: the bounds were divided by the span cell, exactly as the values were. Multiply an axis bound by its span and you get back the range the geometry was measured in.

### C04A

> Compared to plan California (+34 kUSD) and Ohio (+31 kUSD) have the greatest absolute positive variances in net sales

**Tier stack.** One chart per tier, stacked along the page and sharing a category axis. The tiers are separate chart objects positioned to line up, because one chart cannot carry two value axes that both start at zero.

| | |
|---|---|
| Cells used | `A1:Y31` |
| Print area | `'C04A'!$Z$1:$AR$32` |
| Hidden engine columns | `M` |
| Formula cells | 366, in 21 shapes |
| Typed number cells | 38 |

**Typed values** - `C13:D31`. Everything else on the sheet is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in this range, so go by the shading rather than by the count.

**The formulas behind everything else.** Cells that say the same thing about their own position are one row here; the example is the first of them, in ordinary A1.

| Cells | Named | Example | Formula |
|---|---|---|---|
| 57 cells in `P13:U31` | measure scaled, abs_up scaled, abs_dn scaled | `P13` | `=C13/$M$13` |
| `X13:Y31` | rel_up scaled, rel_dn scaled | `X13` | `=I13/$N$13` |
| `E13:E31` | PL | `E13` | `=C13-D13` |
| `F13:F31` | ΔPL% | `F13` | `=IF(E13=0,NA(),D13/E13*100)` |
| `G13:G31` | ΔPL up | `G13` | `=IF($D13>0,$D13,NA())` |
| `H13:H31` | ΔPL down | `H13` | `=IF($D13<0,$D13,NA())` |
| `I13:I31` | ΔPL% up | `I13` | `=IF($F13>0,$F13,NA())` |
| `J13:J31` | ΔPL% down | `J13` | `=IF($F13<0,$F13,NA())` |
| `K13:K31` | up length | `K13` | `=IF($F13>0,$F13,0)` |
| `L13:L31` | down length | `L13` | `=IF($F13<0,-$F13,0)` |
| `O13:O31` | ref scaled | `O13` | `=E13/$M$13` |
| `Q13:Q31` | measure text | `Q13` | `=IF(ISNA(C13),"",TEXT(C13,"0"))` |
| `R13:R31` | var_abs scaled | `R13` | `=D13/$M$13` |
| `S13:S31` | var_abs text | `S13` | `=IF(ISNA(D13),"",TEXT(D13,"+0;-0"))` |
| `V13:V31` | var_rel scaled | `V13` | `=F13/$N$13` |
| `W13:W31` | var_rel text | `W13` | `=IF(ISNA(F13),"",TEXT(F13,"+0;-0"))` |
| `B7` | Subject | `B7` | `=B2&" in "&B3` |
| `B9` | var_abs caption | `B9` | `="Δ"&$B$5&" "&$B$3` |
| `B10` | var_rel caption | `B10` | `="Δ"&$B$5&"%"` |
| `M13` | span unit | `M13` | `=MAX(AGGREGATE(4,6,E13:E31,C13:C31,D13:D31,G13:G31,H13:H31),-AGGREGATE(5,6,E13:E31,C13:C31,D13:D31,G13:G31,H13:H31),0.000000001)` |
| `N13` | span rel | `N13` | `=MAX(AGGREGATE(4,6,F13:F31,I13:I31,J13:J31),-AGGREGATE(5,6,F13:F31,I13:I31,J13:J31),0.000000001)` |

**Chart objects**, in points from the top left of the sheet. The positions are what make separate charts read as one figure, so they are declared rather than dragged.

| Name | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|
| `tier_measure` | 631.5 | 86 | 430 | 400 | 2 | 0 .. 1.0959 |
| `tier_var_abs` | 1066.5 | 86 | 250 | 400 | 3 | -0.4795 .. 0.2055 |
| `tier_var_rel` | 1321.5 | 86 | 170 | 400 | 3 | -1.1124 .. 0.599 |

An axis range that is not round - `0.9` rather than `900` - is the scale block at work: the bounds were divided by the span cell, exactly as the values were. Multiply an axis bound by its span and you get back the range the geometry was measured in.

### C05X

> We expect a plus of 24 kEUR (+15.6%) vs plan until end of the year because of the positive forecast beginning in September

**Tier stack.** One chart per tier, stacked along the page and sharing a category axis. The tiers are separate chart objects positioned to line up, because one chart cannot carry two value axes that both start at zero.

| | |
|---|---|
| Cells used | `A1:AU28` |
| Print area | `'C05X'!$AV$1:$BI$46` |
| Hidden engine columns | `W` |
| Formula cells | 689, in 61 shapes |
| Typed number cells | 25 |

**Typed values** - 25 cells in `D12:E25`. Everything else on the sheet is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in this range, so go by the shading rather than by the count.

**The formulas behind everything else.** Cells that say the same thing about their own position are one row here; the example is the first of them, in ordinary A1.

| Cells | Named | Example | Formula |
|---|---|---|---|
| 51 cells in `AC12:AH28` | bar_ac scaled, wf_base scaled, wf_up_ac scaled | `AC12` | `=J12/$W$12` |
| 40 cells in `G12:T28` | ΔPL%, PY bar, PL bar, AC bar | `G12` | `=NA()` |
| 34 cells in `AE12:AJ28` | bar_fc scaled, wf_dn_ac scaled | `AE12` | `=K12/$W$12` |
| `AT12:AU28` | rel_up_len scaled, rel_dn_len scaled | `AT12` | `=U12/$X$12` |
| 33 cells in `F12:V28` | PL, level, base, up AC | `F12` | `=0` |
| `Y12:Y28` | bar_py scaled | `Y12` | `=H12/$W$12` |
| `Z12:Z28` | bar_py text | `Z12` | `=IF(ISNA(H12),"",TEXT(H12,"#,##0"))` |
| `AA12:AA28` | bar_pl scaled | `AA12` | `=I12/$W$12` |
| `AB12:AB28` | bar_pl text | `AB12` | `=IF(ISNA(I12),"",TEXT(I12,"#,##0"))` |
| `AD12:AD28` | bar_ac text | `AD12` | `=IF(ISNA(J12),"",TEXT(J12,"#,##0"))` |
| `AF12:AF28` | bar_fc text | `AF12` | `=IF(ISNA(K12),"",TEXT(K12,"#,##0"))` |
| `AI12:AI28` | wf_up_ac text | `AI12` | `=IF(ISNA(O12),"",TEXT(O12,"+#,##0"))` |
| `AK12:AK28` | wf_dn_ac text | `AK12` | `=IF(ISNA(P12),"",TEXT(P12,"""-""#,##0"))` |
| `AL12:AL28` | wf_up_fc scaled | `AL12` | `=Q12/$W$12` |
| `AM12:AM28` | wf_up_fc text | `AM12` | `=IF(ISNA(Q12),"",TEXT(Q12,"+#,##0"))` |
| `AN12:AN28` | wf_dn_fc scaled | `AN12` | `=R12/$W$12` |
| `AO12:AO28` | wf_dn_fc text | `AO12` | `=IF(ISNA(R12),"",TEXT(R12,"""-""#,##0"))` |
| `AP12:AP28` | rel_up scaled | `AP12` | `=S12/$X$12` |
| `AQ12:AQ28` | rel_up text | `AQ12` | `=IF(ISNA(S12),"",TEXT(S12,"+0.0;-0.0"))` |
| `AR12:AR28` | rel_dn scaled | `AR12` | `=T12/$X$12` |
| `AS12:AS28` | rel_dn text | `AS12` | `=IF(ISNA(T12),"",TEXT(T12,"+0.0;-0.0"))` |
| `S14:S26` | Jan | `S14` | `=IF(G14>0,G14,NA())` |
| `T14:T26` | Jan | `T14` | `=IF(G14<0,G14,NA())` |
| `U14:U26` | Jan | `U14` | `=IF(G14>0,G14,0)` |
| `V14:V26` | Jan | `V14` | `=IF(G14<0,-G14,0)` |
| `F14:F25` | Jan | `F14` | `=D14-E14` |
| `G14:G25` | Jan | `G14` | `=IF(F14=0,NA(),E14/F14*100)` |
| `I14:I25` | Jan | `I14` | `=F14` |
| `J14:J25` | Jan | `J14` | `=IF($B14="AC",D14,NA())` |
| `K14:K25` | Jan | `K14` | `=IF($B14="FC",D14,NA())` |
| `L14:L25` | Jan | `L14` | `=M14+E14` |
| `M14:M25` | Jan | `M14` | `=L13` |
| `N14:N25` | Jan | `N14` | `=MIN(M14,L14)` |
| `O14:O25` | Jan | `O14` | `=IF(AND($B14="AC",E14>0),E14,0)` |
| `P14:P25` | Jan | `P14` | `=IF(AND($B14="AC",E14<0),-E14,0)` |
| `Q14:Q25` | Jan | `Q14` | `=IF(AND($B14="FC",E14>0),E14,0)` |
| `R14:R25` | Jan | `R14` | `=IF(AND($B14="FC",E14<0),-E14,0)` |
| 3 cells in `M12:M26` | from | `M12` | `=L12` |
| `L27:L28` | vs PL | `L27` | `=$D$26` |
| `O27:O28` | vs PL | `O27` | `=IF(E27>0,E27,0)` |

*21 further named shapes are not listed - this sheet has 61 in all.*

**Chart objects**, in points from the top left of the sheet. The positions are what make separate charts read as one figure, so they are declared rather than dragged.

| Name | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|
| `tier_var_rel` | 1038.5 | 96 | 620 | 130 | 2 | -0.8182 .. 1.0909 |
| `tier_wf` | 1038.5 | 236 | 620 | 110 | 5 | 0.8046 .. 1.1494 |
| `tier_measure` | 1038.5 | 356 | 620 | 330 | 4 | 0 .. 1.1494 |

An axis range that is not round - `0.9` rather than `900` - is the scale block at work: the bounds were divided by the span cell, exactly as the values were. Multiply an axis bound by its span and you get back the range the geometry was measured in.

### C06F

> In Q3 2025, the total decrease in net sales compared to PY was 343 kUSD

**Tier stack.** One chart per tier, stacked along the page and sharing a category axis. The tiers are separate chart objects positioned to line up, because one chart cannot carry two value axes that both start at zero.

| | |
|---|---|
| Cells used | `A1:AM31` |
| Print area | `'C06F'!$AN$1:$BG$35` |
| Hidden engine columns | `T` |
| Formula cells | 644, in 46 shapes |
| Typed number cells | 33 |

**Typed values** - 33 cells in `D12:E30`. Everything else on the sheet is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in this range, so go by the shading rather than by the count.

**The formulas behind everything else.** Cells that say the same thing about their own position are one row here; the example is the first of them, in ordinary A1.

| Cells | Named | Example | Formula |
|---|---|---|---|
| 82 cells in `F12:Q31` | PY, AC bar, PY short, PL bar | `F12` | `=0` |
| 60 cells in `AA12:AF31` | bar_py scaled, wf_base scaled, wf_up scaled | `AA12` | `=K12/$T$12` |
| `X12:Y31` | bar_short scaled, bar_pl scaled | `X12` | `=I12/$T$12` |
| 40 cells in `AC12:AH31` | bar_ac scaled, wf_dn scaled | `AC12` | `=L12/$T$12` |
| `AL12:AM31` | rel_up scaled, rel_dn scaled | `AL12` | `=R12/$U$12` |
| `V12:V31` | bar_state scaled | `V12` | `=H12/$T$12` |
| `W12:W31` | bar_state text | `W12` | `=IF(ISNA(H12),"",TEXT(H12,"#,##0"))` |
| `Z12:Z31` | bar_pl text | `Z12` | `=IF(ISNA(J12),"",TEXT(J12,"#,##0"))` |
| `AB12:AB31` | bar_py text | `AB12` | `=IF(ISNA(K12),"",TEXT(K12,"#,##0"))` |
| `AD12:AD31` | bar_ac text | `AD12` | `=IF(ISNA(L12),"",TEXT(L12,"#,##0"))` |
| `AG12:AG31` | wf_up text | `AG12` | `=IF(ISNA(P12),"",TEXT(P12,"+#,##0"))` |
| `AI12:AI31` | wf_dn text | `AI12` | `=IF(ISNA(Q12),"",TEXT(Q12,"""-""#,##0"))` |
| `AJ12:AJ31` | var_rel scaled | `AJ12` | `=E12/$U$12` |
| `AK12:AK31` | var_rel text | `AK12` | `=IF(ISNA(E12),"",TEXT(E12,"+0;-0"))` |
| 17 cells in `P14:P31` | Iowa | `P14` | `=IF(G14>0,G14,0)` |
| 17 cells in `Q14:Q31` | Iowa | `Q14` | `=IF(G14<0,-G14,0)` |
| `R14:R29` | Iowa | `R14` | `=IF(E14>0,E14,NA())` |
| `S14:S29` | Iowa | `S14` | `=IF(E14<0,E14,NA())` |
| `F14:F28` | Iowa | `F14` | `=IF(E14=-100,NA(),D14/(1+E14/100))` |
| `G14:G28` | Iowa | `G14` | `=D14-F14` |
| `H14:H28` | Iowa | `H14` | `=D14` |
| `I14:I28` | Iowa | `I14` | `=MAX(F14-D14,0)` |
| `M14:M28` | Iowa | `M14` | `=N14+G14` |
| `N14:N28` | Iowa | `N14` | `=M13` |
| `O14:O28` | Iowa | `O14` | `=MIN(N14,M14)` |
| 8 cells in `R12:S31` | ΔPY% up, ΔPY% down | `R12` | `=NA()` |
| 3 cells in `N12:N29` | from | `N12` | `=M12` |
| 2 cells in `M30:N31` | ΔPY | `N30` | `=$D$29` |
| `B7` | Subject | `B7` | `=B2&" in "&B3` |
| `B9` | var_rel caption | `B9` | `="Δ"&$B$5&"%"` |
| `J12` | PL bar | `J12` | `=D12` |
| `T12` | span unit | `T12` | `=MAX(AGGREGATE(4,6,H12:H31,I12:I31,J12:J31,K12:K31,L12:L31,O12:O31,P12:P31,Q12:Q31),-AGGREGATE(5,6,H12:H31,I12:I31,J12:J31,K12:K31,L12:L31,O12:O31,P12:P31,Q12:Q31),0.000000001)` |
| `U12` | span rel | `U12` | `=MAX(AGGREGATE(4,6,E12:E31,R12:R31,S12:S31),-AGGREGATE(5,6,E12:E31,R12:R31,S12:S31),0.000000001)` |
| `D13` | PY | `D13` | `=SUM(F14:F28)` |
| `K13` | PY | `K13` | `=D13` |
| `M13` | PY | `M13` | `=D13` |
| `D29` | AC | `D29` | `=SUM(D14:D28)` |
| `G29` | AC | `G29` | `=D29-$D$13` |
| `L29` | AC | `L29` | `=D29` |
| `M29` | AC | `M29` | `=M28` |

*6 further named shapes are not listed - this sheet has 46 in all.*

**Chart objects**, in points from the top left of the sheet. The positions are what make separate charts read as one figure, so they are declared rather than dragged.

| Name | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|
| `tier_measure` | 970.5 | 86 | 560 | 430 | 5 | 0 .. 1.0255 |
| `tier_wf` | 1540.5 | 86 | 140 | 430 | 3 | 0.7458 .. 1.0255 |
| `tier_var_rel` | 1690.5 | 86 | 200 | 430 | 3 | -1.1111 .. 0.6349 |

An axis range that is not round - `0.9` rather than `900` - is the scale block at work: the bounds were divided by the span cell, exactly as the values were. Multiply an axis bound by its span and you get back the range the geometry was measured in.

### C07C

> Till the end of August our net sales was 192 kUSD lower than planned, because … However, we estimate that full year net sales will be higher than budgeted

**Line.** Series laid out in rows rather than columns, one row per series across the periods, with a combo chart over them.

| | |
|---|---|
| Cells used | `A1:M59` |
| Print area | `'C07C'!$N$1:$AA$36` |
| Formula cells | 183, in 21 shapes |
| Typed number cells | 36 |

**Typed values** - 36 cells in `B11:M17`. Everything else on the sheet is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in this range, so go by the shading rather than by the count.

**The formulas behind everything else.** Cells that say the same thing about their own position are one row here; the example is the first of them, in ordinary A1.

| Cells | Named | Example | Formula |
|---|---|---|---|
| `B32:M38` | PL month scaled | `B32` | `=IF(ISNA(B11),NA(),IF(ISBLANK(B11),NA(),B11/$B$19))` |
| `B54:M59` | AC month text | `B54` | `=IF(ISNA(B12),"",IF(ISBLANK(B12),"",TEXT(B12,"#,##0")))` |
| `B14:B15` | PL cumulative | `B14` | `=SUM(B11:B11)` |
| `C14:C15` | PL cumulative | `C14` | `=SUM(B11:C11)` |
| `D14:D15` | PL cumulative | `D14` | `=SUM(B11:D11)` |
| `E14:E15` | PL cumulative | `E14` | `=SUM(B11:E11)` |
| `F14:F15` | PL cumulative | `F14` | `=SUM(B11:F11)` |
| `G14:G15` | PL cumulative | `G14` | `=SUM(B11:G11)` |
| `H14:H15` | PL cumulative | `H14` | `=SUM(B11:H11)` |
| `I14:I15` | PL cumulative | `I14` | `=SUM(B11:I11)` |
| `B7` | Subject | `B7` | `=B2&IF(B3="",""," in "&B3)` |
| `J14` | PL cumulative | `J14` | `=SUM(B11:J11)` |
| `K14` | PL cumulative | `K14` | `=SUM(B11:K11)` |
| `L14` | PL cumulative | `L14` | `=SUM(B11:L11)` |
| `M14` | PL cumulative | `M14` | `=SUM(B11:M11)` |
| `I16` | FC cumulative | `I16` | `=SUM(B12:I12)+SUM(B13:I13)` |
| `J16` | FC cumulative | `J16` | `=SUM(B12:J12)+SUM(B13:J13)` |
| `K16` | FC cumulative | `K16` | `=SUM(B12:K12)+SUM(B13:K13)` |
| `L16` | FC cumulative | `L16` | `=SUM(B12:L12)+SUM(B13:L13)` |
| `M16` | FC cumulative | `M16` | `=SUM(B12:M12)+SUM(B13:M13)` |
| `B19` | span | `B19` | `=MAX(AGGREGATE(4,6,B11:M11,B12:M12,B13:M13,B14:M14,B15:M15,B16:M16,B17:M17),-AGGREGATE(5,6,B11:M11,B12:M12,B13:M13,B14:M14,B15:M15,B16:M16,B17:M17),0.000000001)` |

**Chart objects**, in points from the top left of the sheet. The positions are what make separate charts read as one figure, so they are declared rather than dragged.

| Name | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|
| `line_chart` | 644 | 81 | 620 | 223.5 | 7 | 0 .. 1.0387 |

An axis range that is not round - `0.9` rather than `900` - is the scale block at work: the bounds were divided by the span cell, exactly as the values were. Multiply an axis bound by its span and you get back the range the geometry was measured in.

### C08H

> Until end of 2025 we plan a slight increase of our raw material stock to 13 tons

**Line.** Series laid out in rows rather than columns, one row per series across the periods, with a combo chart over them.

| | |
|---|---|
| Cells used | `A1:U65` |
| Print area | `'C08H'!$V$1:$AL$81` |
| Formula cells | 262, in 30 shapes |
| Typed number cells | 41 |

**Typed values** - 41 cells in `A12:U15`. Everything else on the sheet is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in this range, so go by the shading rather than by the count.

**The formulas behind everything else.** Cells that say the same thing about their own position are one row here; the example is the first of them, in ordinary A1.

| Cells | Named | Example | Formula |
|---|---|---|---|
| 100 cells in `B35:U42` | Increase scaled | `B35` | `=IF(ISNA(B12),NA(),IF(ISBLANK(B12),NA(),B12/$B$21))` |
| 40 cells in `B58:U65` | Increase text | `B58` | `=IF(ISNA(B12),"",IF(ISBLANK(B12),"",TEXT(B12,"0")))` |
| 38 cells in `B16:U18` | Inventory AC | `N16` | `=NA()` |
| `B14:U14` | Inventory change | `B14` | `=B12-B13` |
| `B19:U19` | Decrease (drawn) | `B19` | `=-B13` |
| `B16:M16` | Inventory AC | `B16` | `=B15` |
| `N18:U18` | Inventory PL | `N18` | `=N15` |
| `M17:N17` | Inventory FC | `M17` | `=M15` |
| `B7` | Subject | `B7` | `=B2&IF(B3="",""," in "&B3)` |
| `B21` | span | `B21` | `=MAX(AGGREGATE(4,6,B16:U16,B17:U17,B18:U18,B12:U12,B19:U19),-AGGREGATE(5,6,B16:U16,B17:U17,B18:U18,B12:U12,B19:U19),0.000000001)` |

20 unnamed one-off formulas: `B15`, `C15`, `D15`, `E15`, `F15`, `G15`, `H15`, `I15`, `J15`, `K15`, `L15`, `M15`, ....

**Chart objects**, in points from the top left of the sheet. The positions are what make separate charts read as one figure, so they are declared rather than dragged.

| Name | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|
| `c08_change` | 865 | 81 | 760 | 150 | 1 | -8 .. 4 |
| `c08_level` | 865 | 241 | 760 | 300 | 3 | -0.3636 .. 1.0909 |
| `c08_flows` | 865 | 241 | 760 | 300 | 2 | -0.3636 .. 1.0909 |

An axis range that is not round - `0.9` rather than `900` - is the scale block at work: the bounds were divided by the span cell, exactly as the values were. Multiply an axis bound by its span and you get back the range the geometry was measured in.

### C09C

> In 2025 we had 45 products of the product line VA in the gross profit segment of 3 mUSD and above

**Xy.** A scatter or bubble plot. These two are the only charts in the library that show a value axis, which is why they are fitted at build time rather than normalised - see the scale block below.

| | |
|---|---|
| Cells used | `A1:E226` |
| Print area | `'C09C'!$F$1:$T$46` |
| Formula cells | 331, in 6 shapes |
| Typed number cells | 362 |

**Typed values** - 362 cells in `B11:E226`. Everything else on the sheet is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in this range, so go by the shading rather than by the count.

**The formulas behind everything else.** Cells that say the same thing about their own position are one row here; the example is the first of them, in ordinary A1.

| Cells | Named | Example | Formula |
|---|---|---|---|
| 149 cells in `E11:E161` | Gross profit in mUSD | `E11` | `=C11*B11/100` |
| `C167:C226` | - | `C167` | `=IF(1*100/$B167>35,NA(),1*100/$B167)` |
| `D167:D226` | - | `D167` | `=IF(2*100/$B167>35,NA(),2*100/$B167)` |
| `E167:E226` | - | `E167` | `=IF(3*100/$B167>35,NA(),3*100/$B167)` |
| `B7` | Subject | `B7` | `=B2&IF(B3="",""," in "&B3)` |
| `E163` | VA at 3 mUSD or more | `E163` | `=SUMPRODUCT((D11:D161="VA")*(E11:E161>=3))` |

**Chart objects**, in points from the top left of the sheet. The positions are what make separate charts read as one figure, so they are declared rather than dragged.

| Name | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|
| `xy_C09C` | 310.5 | 99 | 700 | 560 | 6 | 0 .. 35 |

An axis range that is not round - `0.9` rather than `900` - is the scale block at work: the bounds were divided by the span cell, exactly as the values were. Multiply an axis bound by its span and you get back the range the geometry was measured in.

### C10D

> The two SBUs acquired in December 2025 are positioned in little attractive markets

**Xy.** A scatter or bubble plot. These two are the only charts in the library that show a value axis, which is why they are fitted at build time rather than normalised - see the scale block below.

| | |
|---|---|
| Cells used | `A1:E32` |
| Print area | `'C10D'!$F$1:$T$40` |
| Formula cells | 2, in 2 shapes |
| Typed number cells | 60 |

**Typed values** - 60 cells in `B11:E32`. Everything else on the sheet is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in this range, so go by the shading rather than by the count.

**The formulas behind everything else.** Cells that say the same thing about their own position are one row here; the example is the first of them, in ordinary A1.

| Cells | Named | Example | Formula |
|---|---|---|---|
| `B7` | Subject | `B7` | `=B2&IF(B3="",""," in "&B3)` |
| `E10` | PY | `E10` | `=B8&" in "&C8` |

**Chart objects**, in points from the top left of the sheet. The positions are what make separate charts read as one figure, so they are declared rather than dragged.

| Name | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|
| `xy_C10D` | 321.5 | 99 | 660 | 470 | 3 | 0 .. 1.25 |

An axis range that is not round - `0.9` rather than `900` - is the scale block at work: the bounds were divided by the span cell, exactly as the values were. Multiply an axis bound by its span and you get back the range the geometry was measured in.

### C11A

> We plan to achieve a ROI of 19,2% in 2024 despite of increasing invested capital

**Tree.** Boxes and connectors: several small charts and the arithmetic between them, with boxes that share a unit sharing a scale.

| | |
|---|---|
| Cells used | `A1:H62` |
| Print area | `'C11A'!$I$4:$Z$88` |
| Formula cells | 123, in 15 shapes |
| Typed number cells | 21 |

**Typed values** - `B13:H15`. Everything else on the sheet is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in this range, so go by the shading rather than by the count.

**The formulas behind everything else.** Cells that say the same thing about their own position are one row here; the example is the first of them, in ordinary A1.

| Cells | Named | Example | Formula |
|---|---|---|---|
| `B57:H62` | return text | `B57` | `=IF(ISNA(B13),"",IF(ISBLANK(B13),"",TEXT(B13,"0.0")))` |
| `B35:H37` | return scaled | `B35` | `=IF(ISNA(B13),NA(),IF(ISBLANK(B13),NA(),B13/$B$22))` |
| 14 cells in `B38:H40` | ros scaled | `B38` | `=IF(ISNA(B16),NA(),IF(ISBLANK(B16),NA(),B16/$B$20))` |
| `B16:H16` | Return on sales | `B16` | `=B13/B14*100` |
| `B17:H17` | Capital turnover | `B17` | `=B14/B15` |
| `B18:H18` | Return on investment | `B18` | `=B13/B15*100` |
| `B39:H39` | turnover scaled | `B39` | `=IF(ISNA(B17),NA(),IF(ISBLANK(B17),NA(),B17/$B$21))` |
| `C11:H11` | AC, PL | `C11` | `=IF(C10=B10,"",C10)` |
| 5 cells in `B12:H12` | Axis year | `B12` | `=B9` |
| `C12:D12` | Axis year | `C12` | `=""` |
| `B7` | Subject | `B7` | `=B2&IF(B3="",""," in "&B3)` |
| `B11` | AC | `B11` | `=B10` |
| `B20` | span percent | `B20` | `=MAX(AGGREGATE(4,6,B18:H18,B16:H16),-AGGREGATE(5,6,B18:H18,B16:H16),0.000000001)` |
| `B21` | span turnover | `B21` | `=MAX(AGGREGATE(4,6,B17:H17),-AGGREGATE(5,6,B17:H17),0.000000001)` |
| `B22` | span kEUR | `B22` | `=MAX(AGGREGATE(4,6,B13:H13,B14:H14,B15:H15),-AGGREGATE(5,6,B13:H13,B14:H14,B15:H15),0.000000001)` |

**Chart objects**, in points from the top left of the sheet. The positions are what make separate charts read as one figure, so they are declared rather than dragged.

| Name | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|
| `tree_roi` | 441.5 | 243.2 | 232 | 302.2 | 1 | -1.005 .. 1.407 |
| `tree_ros` | 715.5 | 115 | 232 | 244.1 | 1 | -0.804 .. 1.005 |
| `tree_turnover` | 715.5 | 362.9 | 232 | 308.5 | 1 | -0.3599 .. 1.4397 |
| `tree_return` | 989.5 | 163 | 232 | 94.1 | 1 | -0.1805 .. 0.2708 |
| `tree_net_sales` | 989.5 | 266.1 | 232 | 180 | 1 | -0.1805 .. 1.1733 |
| `tree_capital` | 989.5 | 452.3 | 232 | 171.6 | 1 | -0.1805 .. 1.083 |

An axis range that is not round - `0.9` rather than `900` - is the scale block at work: the bounds were divided by the span cell, exactly as the values were. Multiply an axis bound by its span and you get back the range the geometry was measured in.

### C12A

> Compared to 2024, the higher operating expenses (+187 kEUR) were mainly compensated by higher license sales (+183 kEUR), leading to a higher group result (+91 kEUR)

**Tier stack.** One chart per tier, stacked along the page and sharing a category axis. The tiers are separate chart objects positioned to line up, because one chart cannot carry two value axes that both start at zero.

| | |
|---|---|
| Cells used | `A1:AU32` |
| Print area | `'C12A'!$AV$1:$BS$33` |
| Hidden engine columns | `AP`, `AU`, `Y` |
| Formula cells | 861, in 46 shapes |
| Typed number cells | 48 |

**Typed values** - 48 cells in `B13:D32`. Everything else on the sheet is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in this range, so go by the shading rather than by the count.

**The formulas behind everything else.** Cells that say the same thing about their own position are one row here; the example is the first of them, in ordinary A1.

| Cells | Named | Example | Formula |
|---|---|---|---|
| 80 cells in `Z13:AF32` | py_base scaled, py_add scaled, ac_base scaled, ac_add scaled | `Z13` | `=I13/$Y$13` |
| `AL13:AO32` | abs_pos_good scaled, abs_pos_bad scaled, abs_neg_good scaled, abs_neg_bad scaled | `AL13` | `=Q13/$Y$13` |
| `AQ13:AT32` | rel_pos_good clipped, rel_pos_bad clipped, rel_neg_good clipped, rel_neg_bad clipped | `AQ13` | `=IF(ISNA(U13),NA(),MEDIAN(-31.5,U13,238.5))` |
| 40 cells in `I13:N32` | PY base, AC base | `I13` | `=MIN(H13,G13)` |
| 40 cells in `AB13:AG32` | py_add text, ac_add text | `AB13` | `=IF(ISNA(J13),"",TEXT(J13,"#,##0"))` |
| 40 cells in `AC13:AH32` | py_sub scaled, ac_sub scaled | `AC13` | `=K13/$Y$13` |
| 40 cells in `AD13:AI32` | py_sub text, ac_sub text | `AD13` | `=IF(ISNA(K13),"",TEXT(K13,"#,##0"))` |
| 26 cells in `H14:M31` | + Consulting | `H14` | `=G13` |
| `E13:E32` | ΔPY | `E13` | `=D13-C13` |
| `F13:F32` | ΔPY% | `F13` | `=IF(C13=0,NA(),(D13-C13)/C13*100)` |
| `J13:J32` | PY adds | `J13` | `=IF(B13>0,ABS(G13-H13),0)` |
| `K13:K32` | PY subs | `K13` | `=IF(B13<0,ABS(G13-H13),0)` |
| `O13:O32` | AC adds | `O13` | `=IF(B13>0,ABS(L13-M13),0)` |
| `P13:P32` | AC subs | `P13` | `=IF(B13<0,ABS(L13-M13),0)` |
| `Q13:Q32` | ΔPY up ok | `Q13` | `=IF(AND($E13>0,$B13>0),$E13,NA())` |
| `R13:R32` | ΔPY up bad | `R13` | `=IF(AND($E13>0,$B13<0),$E13,NA())` |
| `S13:S32` | ΔPY dn ok | `S13` | `=IF(AND($E13<0,$B13<0),$E13,NA())` |
| `T13:T32` | ΔPY dn bad | `T13` | `=IF(AND($E13<0,$B13>0),$E13,NA())` |
| `U13:U32` | ΔPY% up ok | `U13` | `=IF(AND($F13>0,$B13>0),$F13,NA())` |
| `V13:V32` | ΔPY% up bad | `V13` | `=IF(AND($F13>0,$B13<0),$F13,NA())` |
| `W13:W32` | ΔPY% dn ok | `W13` | `=IF(AND($F13<0,$B13<0),$F13,NA())` |
| `X13:X32` | ΔPY% dn bad | `X13` | `=IF(AND($F13<0,$B13>0),$F13,NA())` |
| `AJ13:AJ32` | var_abs scaled | `AJ13` | `=E13/$Y$13` |
| `AK13:AK32` | var_abs text | `AK13` | `=IF(ISNA(E13),"",TEXT(E13,"+0;-0"))` |
| `AP13:AP32` | var_rel clipped | `AP13` | `=IF(ISNA(F13),NA(),MEDIAN(-31.5,F13,238.5))` |
| `AU13:AU32` | ΔPY% text | `AU13` | `=IF(ISNA(F13),"",IF(F13>265,TEXT(F13,"+0;-0")&REPT(UNICHAR(9658),MIN(3,MAX(1,INT(F13/265)))),IF(F13<-35,REPT(UNICHAR(9668),MIN(3,MAX(1,INT(-F13/35))))&TEXT(F13,"+0;-0"),TEXT(F13,"+0;-0"))))` |
| 13 cells in `G14:G31` | + Consulting | `G14` | `=G13+B14*C14` |
| 13 cells in `L14:L31` | + Consulting | `L14` | `=L13+B14*D14` |
| 12 cells in `H13:M32` | PY from, AC from | `H13` | `=0` |
| 12 cells in `G17:L32` | - | `G17` | `=G16` |
| 5 cells in `C17:C32` | - | `C17` | `=G16` |
| 5 cells in `D17:D32` | - | `D17` | `=L16` |
| 2 cells in `H24:M24` | - Operating expenses | `H24` | `=G18` |
| `B7` | Subject | `B7` | `=B2&" in "&B3` |
| `B9` | var_abs caption | `B9` | `="Δ"&$B$5&" "&$B$3` |
| `B10` | var_rel caption | `B10` | `="Δ"&$B$5&"%"` |
| `G13` | PY level | `G13` | `=0+B13*C13` |
| `L13` | AC level | `L13` | `=0+B13*D13` |
| `Y13` | span unit | `Y13` | `=MAX(AGGREGATE(4,6,I13:I32,J13:J32,K13:K32,N13:N32,O13:O32,P13:P32,E13:E32,Q13:Q32,R13:R32,S13:S32,T13:T32),-AGGREGATE(5,6,I13:I32,J13:J32,K13:K32,N13:N32,O13:O32,P13:P32,E13:E32,Q13:Q32,R13:R32,S13:S32,T13:T32),0.000000001)` |
| `A17` | + Other revenue | `A17` | `= Sales revenue` |

*6 further named shapes are not listed - this sheet has 46 in all.*

**Chart objects**, in points from the top left of the sheet. The positions are what make separate charts read as one figure, so they are declared rather than dragged.

| Name | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|
| `tier_wf_py` | 1177.5 | 86 | 400 | 402 | 3 | 0 .. 1.0556 |
| `tier_wf_ac` | 1587.5 | 86 | 300 | 402 | 3 | 0 .. 1.0556 |
| `tier_var_abs` | 1897.5 | 86 | 175 | 402 | 5 | -0.0613 .. 0.4288 |
| `tier_var_rel` | 2082.5 | 86 | 200 | 402 | 5 | -35 .. 265 |

An axis range that is not round - `0.9` rather than `900` - is the scale block at work: the bounds were divided by the span cell, exactly as the values were. Multiply an axis bound by its span and you get back the range the geometry was measured in.

### C13D

> Berlin will further be above overall average in 2027 and 2028, in 2028 its profits will be 293% above location average

**Panel.** A grid of small charts drawn as one native chart object by the `panel-charts` skill, plus a reference panel drawn beside it.

| | |
|---|---|
| Cells used | `A1:BU106` |
| Print area | `'C13D'!$A$109:$N$148` |
| Formula cells | 2124, in 101 shapes |
| Typed number cells | 415 |

**Typed values** - 415 cells in `B11:BS106`. Everything else on the sheet is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in this range, so go by the shading rather than by the count.

**The formulas behind everything else.** Cells that say the same thing about their own position are one row here; the example is the first of them, in ordinary A1.

| Cells | Named | Example | Formula |
|---|---|---|---|
| `Q53:Q106` | pos | `Q53` | `=P53-$Q$41` |
| `R53:R106` | blk | `R53` | `=IF(Q53<=0,-1,INT((Q53-1)/($Q$39+$Q$40)))` |
| `S53:S106` | sub | `S53` | `=IF(Q53<=0,-1,MOD(Q53-1,($Q$39+$Q$40)))` |
| `T53:T106` | k | `T53` | `=IF(OR(S53<0,S53>=$Q$39),-1,INT(S53/$Q$38))` |
| `U53:U106` | e | `U53` | `=IF(OR(S53<0,S53>=$Q$39),-1,MOD(S53,$Q$38))` |
| `V53:V106` | sep | `V53` | `=IF(OR(Q53<=0,S53<0,S53>=$Q$39,R53>=$Q$36),1,0)` |
| `W53:W106` | kx | `W53` | `=IF(V53=1,-1,T53)` |
| `X53:X106` | blkx | `X53` | `=IF(V53=1,-1,R53)` |
| `Y53:Y106` | cat | `Y53` | `=IF(V53=1,"",IF(AND(MOD(W53,3)=0,U53=INT(($Q$38-1)/2)),INDEX($B$34:$M$34,W53+1),""))` |
| `Z53:Z106` | ΔØ % b0 + | `Z53` | `=IF(V53=1,NA(),IF(INDEX($B$35:$M$50,($Q$32-1-0)*$Q$33+X53+1,0*$Q$34+W53+1)="",NA(),IF(INDEX($B$35:$M$50,($Q$32-1-0)*$Q$33+X53+1,0*$Q$34+W53+1)>=0,0+(INDEX($B$35:$M$50,($Q$32-1-0)*$Q$33+X53+1,0*$Q$34+W53+1)-$Q$46)/$Q$48*$Q$42,NA())))` |
| `AA53:AA106` | ΔØ % b0 - | `AA53` | `=IF(V53=1,NA(),IF(INDEX($B$35:$M$50,($Q$32-1-0)*$Q$33+X53+1,0*$Q$34+W53+1)="",NA(),IF(INDEX($B$35:$M$50,($Q$32-1-0)*$Q$33+X53+1,0*$Q$34+W53+1)<0,0+(INDEX($B$35:$M$50,($Q$32-1-0)*$Q$33+X53+1,0*$Q$34+W53+1)-$Q$46)/$Q$48*$Q$42,NA())))` |
| `AB53:AB106` | ΔØ % b1 + | `AB53` | `=IF(V53=1,NA(),IF(INDEX($B$35:$M$50,($Q$32-1-1)*$Q$33+X53+1,0*$Q$34+W53+1)="",NA(),IF(INDEX($B$35:$M$50,($Q$32-1-1)*$Q$33+X53+1,0*$Q$34+W53+1)>=0,1+(INDEX($B$35:$M$50,($Q$32-1-1)*$Q$33+X53+1,0*$Q$34+W53+1)-$Q$46)/$Q$48*$Q$42,NA())))` |
| `AC53:AC106` | ΔØ % b1 - | `AC53` | `=IF(V53=1,NA(),IF(INDEX($B$35:$M$50,($Q$32-1-1)*$Q$33+X53+1,0*$Q$34+W53+1)="",NA(),IF(INDEX($B$35:$M$50,($Q$32-1-1)*$Q$33+X53+1,0*$Q$34+W53+1)<0,1+(INDEX($B$35:$M$50,($Q$32-1-1)*$Q$33+X53+1,0*$Q$34+W53+1)-$Q$46)/$Q$48*$Q$42,NA())))` |
| `AD53:AD106` | ΔØ % b2 + | `AD53` | `=IF(V53=1,NA(),IF(INDEX($B$35:$M$50,($Q$32-1-2)*$Q$33+X53+1,0*$Q$34+W53+1)="",NA(),IF(INDEX($B$35:$M$50,($Q$32-1-2)*$Q$33+X53+1,0*$Q$34+W53+1)>=0,2+(INDEX($B$35:$M$50,($Q$32-1-2)*$Q$33+X53+1,0*$Q$34+W53+1)-$Q$46)/$Q$48*$Q$42,NA())))` |
| `AE53:AE106` | ΔØ % b2 - | `AE53` | `=IF(V53=1,NA(),IF(INDEX($B$35:$M$50,($Q$32-1-2)*$Q$33+X53+1,0*$Q$34+W53+1)="",NA(),IF(INDEX($B$35:$M$50,($Q$32-1-2)*$Q$33+X53+1,0*$Q$34+W53+1)<0,2+(INDEX($B$35:$M$50,($Q$32-1-2)*$Q$33+X53+1,0*$Q$34+W53+1)-$Q$46)/$Q$48*$Q$42,NA())))` |
| `AF53:AF106` | ΔØ % b3 + | `AF53` | `=IF(V53=1,NA(),IF(INDEX($B$35:$M$50,($Q$32-1-3)*$Q$33+X53+1,0*$Q$34+W53+1)="",NA(),IF(INDEX($B$35:$M$50,($Q$32-1-3)*$Q$33+X53+1,0*$Q$34+W53+1)>=0,3+(INDEX($B$35:$M$50,($Q$32-1-3)*$Q$33+X53+1,0*$Q$34+W53+1)-$Q$46)/$Q$48*$Q$42,NA())))` |
| `AG53:AG106` | ΔØ % b3 - | `AG53` | `=IF(V53=1,NA(),IF(INDEX($B$35:$M$50,($Q$32-1-3)*$Q$33+X53+1,0*$Q$34+W53+1)="",NA(),IF(INDEX($B$35:$M$50,($Q$32-1-3)*$Q$33+X53+1,0*$Q$34+W53+1)<0,3+(INDEX($B$35:$M$50,($Q$32-1-3)*$Q$33+X53+1,0*$Q$34+W53+1)-$Q$46)/$Q$48*$Q$42,NA())))` |
| `AI53:AI106` | stem ΔØ % b0 + | `AI53` | `=IF(V53=1,0,IF(INDEX($B$35:$M$50,($Q$32-1-0)*$Q$33+X53+1,0*$Q$34+W53+1)="",0,IF(INDEX($B$35:$M$50,($Q$32-1-0)*$Q$33+X53+1,0*$Q$34+W53+1)>=0,ABS(INDEX($B$35:$M$50,($Q$32-1-0)*$Q$33+X53+1,0*$Q$34+W53+1))/$Q$48*$Q$42,0)))` |
| `AJ53:AJ106` | stem ΔØ % b0 - | `AJ53` | `=IF(V53=1,0,IF(INDEX($B$35:$M$50,($Q$32-1-0)*$Q$33+X53+1,0*$Q$34+W53+1)="",0,IF(INDEX($B$35:$M$50,($Q$32-1-0)*$Q$33+X53+1,0*$Q$34+W53+1)<0,ABS(INDEX($B$35:$M$50,($Q$32-1-0)*$Q$33+X53+1,0*$Q$34+W53+1))/$Q$48*$Q$42,0)))` |
| `AK53:AK106` | stem ΔØ % b1 + | `AK53` | `=IF(V53=1,0,IF(INDEX($B$35:$M$50,($Q$32-1-1)*$Q$33+X53+1,0*$Q$34+W53+1)="",0,IF(INDEX($B$35:$M$50,($Q$32-1-1)*$Q$33+X53+1,0*$Q$34+W53+1)>=0,ABS(INDEX($B$35:$M$50,($Q$32-1-1)*$Q$33+X53+1,0*$Q$34+W53+1))/$Q$48*$Q$42,0)))` |
| `AL53:AL106` | stem ΔØ % b1 - | `AL53` | `=IF(V53=1,0,IF(INDEX($B$35:$M$50,($Q$32-1-1)*$Q$33+X53+1,0*$Q$34+W53+1)="",0,IF(INDEX($B$35:$M$50,($Q$32-1-1)*$Q$33+X53+1,0*$Q$34+W53+1)<0,ABS(INDEX($B$35:$M$50,($Q$32-1-1)*$Q$33+X53+1,0*$Q$34+W53+1))/$Q$48*$Q$42,0)))` |
| `AM53:AM106` | stem ΔØ % b2 + | `AM53` | `=IF(V53=1,0,IF(INDEX($B$35:$M$50,($Q$32-1-2)*$Q$33+X53+1,0*$Q$34+W53+1)="",0,IF(INDEX($B$35:$M$50,($Q$32-1-2)*$Q$33+X53+1,0*$Q$34+W53+1)>=0,ABS(INDEX($B$35:$M$50,($Q$32-1-2)*$Q$33+X53+1,0*$Q$34+W53+1))/$Q$48*$Q$42,0)))` |
| `AN53:AN106` | stem ΔØ % b2 - | `AN53` | `=IF(V53=1,0,IF(INDEX($B$35:$M$50,($Q$32-1-2)*$Q$33+X53+1,0*$Q$34+W53+1)="",0,IF(INDEX($B$35:$M$50,($Q$32-1-2)*$Q$33+X53+1,0*$Q$34+W53+1)<0,ABS(INDEX($B$35:$M$50,($Q$32-1-2)*$Q$33+X53+1,0*$Q$34+W53+1))/$Q$48*$Q$42,0)))` |
| `AO53:AO106` | stem ΔØ % b3 + | `AO53` | `=IF(V53=1,0,IF(INDEX($B$35:$M$50,($Q$32-1-3)*$Q$33+X53+1,0*$Q$34+W53+1)="",0,IF(INDEX($B$35:$M$50,($Q$32-1-3)*$Q$33+X53+1,0*$Q$34+W53+1)>=0,ABS(INDEX($B$35:$M$50,($Q$32-1-3)*$Q$33+X53+1,0*$Q$34+W53+1))/$Q$48*$Q$42,0)))` |
| `AP53:AP106` | stem ΔØ % b3 - | `AP53` | `=IF(V53=1,0,IF(INDEX($B$35:$M$50,($Q$32-1-3)*$Q$33+X53+1,0*$Q$34+W53+1)="",0,IF(INDEX($B$35:$M$50,($Q$32-1-3)*$Q$33+X53+1,0*$Q$34+W53+1)<0,ABS(INDEX($B$35:$M$50,($Q$32-1-3)*$Q$33+X53+1,0*$Q$34+W53+1))/$Q$48*$Q$42,0)))` |
| `AR53:AR106` | label ΔØ % b0 + | `AR53` | `=IF(V53=1,"",IF(INDEX($B$35:$M$50,($Q$32-1-0)*$Q$33+X53+1,0*$Q$34+W53+1)="","",IF(INDEX($B$35:$M$50,($Q$32-1-0)*$Q$33+X53+1,0*$Q$34+W53+1)>=0,TEXT(INDEX($B$35:$M$50,($Q$32-1-0)*$Q$33+X53+1,0*$Q$34+W53+1),"+0;-0;0"),"")))` |
| `AS53:AS106` | label ΔØ % b0 - | `AS53` | `=IF(V53=1,"",IF(INDEX($B$35:$M$50,($Q$32-1-0)*$Q$33+X53+1,0*$Q$34+W53+1)="","",IF(INDEX($B$35:$M$50,($Q$32-1-0)*$Q$33+X53+1,0*$Q$34+W53+1)<0,TEXT(INDEX($B$35:$M$50,($Q$32-1-0)*$Q$33+X53+1,0*$Q$34+W53+1),"+0;-0;0"),"")))` |
| `AT53:AT106` | label ΔØ % b1 + | `AT53` | `=IF(V53=1,"",IF(INDEX($B$35:$M$50,($Q$32-1-1)*$Q$33+X53+1,0*$Q$34+W53+1)="","",IF(INDEX($B$35:$M$50,($Q$32-1-1)*$Q$33+X53+1,0*$Q$34+W53+1)>=0,TEXT(INDEX($B$35:$M$50,($Q$32-1-1)*$Q$33+X53+1,0*$Q$34+W53+1),"+0;-0;0"),"")))` |
| `AU53:AU106` | label ΔØ % b1 - | `AU53` | `=IF(V53=1,"",IF(INDEX($B$35:$M$50,($Q$32-1-1)*$Q$33+X53+1,0*$Q$34+W53+1)="","",IF(INDEX($B$35:$M$50,($Q$32-1-1)*$Q$33+X53+1,0*$Q$34+W53+1)<0,TEXT(INDEX($B$35:$M$50,($Q$32-1-1)*$Q$33+X53+1,0*$Q$34+W53+1),"+0;-0;0"),"")))` |
| `AV53:AV106` | label ΔØ % b2 + | `AV53` | `=IF(V53=1,"",IF(INDEX($B$35:$M$50,($Q$32-1-2)*$Q$33+X53+1,0*$Q$34+W53+1)="","",IF(INDEX($B$35:$M$50,($Q$32-1-2)*$Q$33+X53+1,0*$Q$34+W53+1)>=0,TEXT(INDEX($B$35:$M$50,($Q$32-1-2)*$Q$33+X53+1,0*$Q$34+W53+1),"+0;-0;0"),"")))` |
| `AW53:AW106` | label ΔØ % b2 - | `AW53` | `=IF(V53=1,"",IF(INDEX($B$35:$M$50,($Q$32-1-2)*$Q$33+X53+1,0*$Q$34+W53+1)="","",IF(INDEX($B$35:$M$50,($Q$32-1-2)*$Q$33+X53+1,0*$Q$34+W53+1)<0,TEXT(INDEX($B$35:$M$50,($Q$32-1-2)*$Q$33+X53+1,0*$Q$34+W53+1),"+0;-0;0"),"")))` |
| `AX53:AX106` | label ΔØ % b3 + | `AX53` | `=IF(V53=1,"",IF(INDEX($B$35:$M$50,($Q$32-1-3)*$Q$33+X53+1,0*$Q$34+W53+1)="","",IF(INDEX($B$35:$M$50,($Q$32-1-3)*$Q$33+X53+1,0*$Q$34+W53+1)>=0,TEXT(INDEX($B$35:$M$50,($Q$32-1-3)*$Q$33+X53+1,0*$Q$34+W53+1),"+0;-0;0"),"")))` |
| `AY53:AY106` | label ΔØ % b3 - | `AY53` | `=IF(V53=1,"",IF(INDEX($B$35:$M$50,($Q$32-1-3)*$Q$33+X53+1,0*$Q$34+W53+1)="","",IF(INDEX($B$35:$M$50,($Q$32-1-3)*$Q$33+X53+1,0*$Q$34+W53+1)<0,TEXT(INDEX($B$35:$M$50,($Q$32-1-3)*$Q$33+X53+1,0*$Q$34+W53+1),"+0;-0;0"),"")))` |
| 38 cells in `BB55:BS100` | - | `BB55` | `=NA()` |
| `B26:M26` | average 25 locations | `B26` | `=AVERAGE(B11:B25)` |
| `B35:M35` | 17, 18, 19, 20 | `B35` | `=(B11-B26)/B26*100` |
| `B36:M36` | Frankfurt | `B36` | `=(B12-B26)/B26*100` |
| `B37:M37` | Munich | `B37` | `=(B13-B26)/B26*100` |
| `B38:M38` | Berlin | `B38` | `=(B14-B26)/B26*100` |
| `B39:M39` | Zurich | `B39` | `=(B15-B26)/B26*100` |

*33 further named shapes are not listed - this sheet has 73 in all.*

28 unnamed one-off formulas: `BL53`, `BO53`, `BL54`, `BO54`, `BL55`, `BO55`, `BL56`, `BO56`, `BL57`, `BO57`, `BL58`, `BO58`, ....

**Chart objects**, in points from the top left of the sheet. The positions are what make separate charts read as one figure, so they are declared rather than dragged.

| Name | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|
| `panel_grid` | 0 | 1637 | 720 | 500 | 14 | 0 .. 4 |
| `reference_panel` | 549.9 | 2015.8 | 155.1 | 95.9 | 1 | 0 .. 114.743 |

An axis range that is not round - `0.9` rather than `900` - is the scale block at work: the bounds were divided by the span cell, exactly as the values were. Multiply an axis bound by its span and you get back the range the geometry was measured in.

### T01B

**Table.** A printed grid, with some columns *drawn* instead of printed - the drawn ones are ordinary variance charts sized to the row pitch and positioned against the rows they belong to.

| | |
|---|---|
| Cells used | `A1:O35` |
| Print area | `'T01B'!$A$9:$O$35` |
| Formula cells | 190, in 13 shapes |
| Typed number cells | 104 |

**Typed values** - 104 cells in `A15:O35`. Everything else on the sheet is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in this range, so go by the shading rather than by the count.

**The formulas behind everything else.** Cells that say the same thing about their own position are one row here; the example is the first of them, in ordinary A1.

| Cells | Named | Example | Formula |
|---|---|---|---|
| 40 cells in `D15:L34` | ΔPY | `D15` | `=C15-A15` |
| 40 cells in `E15:M34` | ΔPY% | `E15` | `=IF(A15=0,NA(),(C15-A15)/A15*100)` |
| 40 cells in `F15:N34` | ΔPL | `F15` | `=C15-B15` |
| 40 cells in `G15:O34` | ΔPL% | `G15` | `=IF(B15=0,NA(),(C15-B15)/B15*100)` |
| 12 cells in `A28:K33` | - | `A28` | `=ABS(SUM(A24:A27))` |
| 6 cells in `A23:K23` | - | `A23` | `=ABS(SUM(A15:A22))` |
| 6 cells in `A34:K34` | - | `A34` | `=SUM(A15:A22)+SUM(A24:A27)+SUM(A29:A32)` |
| `B7` | Subject | `B7` | `=B2&IF(B3="",""," in "&B3)` |

5 unnamed one-off formulas: `A9`, `A10`, `A11`, `A13`, `I13`.

### T02A

**Table.** A printed grid, with some columns *drawn* instead of printed - the drawn ones are ordinary variance charts sized to the row pitch and positioned against the rows they belong to.

| | |
|---|---|
| Cells used | `A1:AB34` |
| Print area | `'T02A'!$A$9:$J$35` |
| Formula cells | 424, in 28 shapes |
| Typed number cells | 64 |

**Typed values** - 64 cells in `A15:G32`. Everything else on the sheet is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in this range, so go by the shading rather than by the count.

**The formulas behind everything else.** Cells that say the same thing about their own position are one row here; the example is the first of them, in ordinary A1.

| Cells | Named | Example | Formula |
|---|---|---|---|
| 40 cells in `C15:H34` | ΔPL | `C15` | `=B15-A15` |
| 40 cells in `D15:I34` | ΔPL% | `D15` | `=IF(A15=0,NA(),(B15-A15)/A15*100)` |
| `M15:M34` | ΔPL scaled | `M15` | `=IF(ISBLANK(C15),NA(),C15/$K$15)` |
| `N15:N34` | ΔPL text | `N15` | `=IF(ISBLANK(C15),"",TEXT(C15,"[>0.5]+0;[<-0.5]-0;0"))` |
| `O15:O34` | ΔPL good | `O15` | `=IF(ISBLANK(C15),NA(),IF(C15>0,C15/$K$15,NA()))` |
| `P15:P34` | ΔPL bad | `P15` | `=IF(ISBLANK(C15),NA(),IF(C15<0,C15/$K$15,NA()))` |
| `Q15:Q34` | ΔPL% scaled | `Q15` | `=IF(ISBLANK(D15),NA(),D15/$L$15)` |
| `R15:R34` | ΔPL% text | `R15` | `=IF(ISBLANK(D15),"",TEXT(D15,"[>0.05]+0.0""%"";[<-0.05]-0.0""%"";0.0""%"""))` |
| `S15:S34` | ΔPL% good | `S15` | `=IF(ISBLANK(D15),NA(),IF(D15>0,D15/$L$15,NA()))` |
| `T15:T34` | ΔPL% bad | `T15` | `=IF(ISBLANK(D15),NA(),IF(D15<0,D15/$L$15,NA()))` |
| `U15:U34` | ΔPL scaled | `U15` | `=IF(ISBLANK(H15),NA(),H15/$K$15)` |
| `V15:V34` | ΔPL text | `V15` | `=IF(ISBLANK(H15),"",TEXT(H15,"[>0.5]+0;[<-0.5]-0;0"))` |
| `W15:W34` | ΔPL good | `W15` | `=IF(ISBLANK(H15),NA(),IF(H15>0,H15/$K$15,NA()))` |
| `X15:X34` | ΔPL bad | `X15` | `=IF(ISBLANK(H15),NA(),IF(H15<0,H15/$K$15,NA()))` |
| `Y15:Y34` | ΔPL% scaled | `Y15` | `=IF(ISBLANK(I15),NA(),I15/$L$15)` |
| `Z15:Z34` | ΔPL% text | `Z15` | `=IF(ISBLANK(I15),"",TEXT(I15,"[>0.05]+0.0""%"";[<-0.05]-0.0""%"";0.0""%"""))` |
| `AA15:AA34` | ΔPL% good | `AA15` | `=IF(ISBLANK(I15),NA(),IF(I15>0,I15/$L$15,NA()))` |
| `AB15:AB34` | ΔPL% bad | `AB15` | `=IF(ISBLANK(I15),NA(),IF(I15<0,I15/$L$15,NA()))` |
| 8 cells in `A28:G33` | - | `A28` | `=ABS(SUM(A24:A27))` |
| 4 cells in `A23:G23` | - | `A23` | `=ABS(SUM(A15:A22))` |
| 4 cells in `A34:G34` | - | `A34` | `=SUM(A15:A22)+SUM(A24:A27)+SUM(A29:A32)` |
| `K15:L15` | span unit, span rel | `K15` | `=MAX(AGGREGATE(4,6,C15:C34,H15:H34),-AGGREGATE(5,6,C15:C34,H15:H34),0.000000001)` |
| `B7` | Subject | `B7` | `=B2&IF(B3="",""," in "&B3)` |

5 unnamed one-off formulas: `A9`, `A10`, `A11`, `A13`, `F13`.

**Chart objects**, in points from the top left of the sheet. The positions are what make separate charts read as one figure, so they are declared rather than dragged.

| Name | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|
| `panel_dpl_november` | 113.5 | 205.2 | 96.7 | 303.9 | 3 | -0.7101 .. 0.4769 |
| `panel_dplp_november` | 206.1 | 205.1 | 186.4 | 303.9 | 3 | -1.5714 .. 1.6286 |
| `panel_dpl_ytd_november` | 598.6 | 205.1 | 186.4 | 303.9 | 3 | -1.3989 .. 0.975 |
| `panel_dplp_ytd_november` | 781.1 | 205.1 | 186.4 | 303.9 | 3 | -1.5714 .. 1.6286 |

An axis range that is not round - `0.9` rather than `900` - is the scale block at work: the bounds were divided by the span cell, exactly as the values were. Multiply an axis bound by its span and you get back the range the geometry was measured in.

### T03A

**Table.** A printed grid, with some columns *drawn* instead of printed - the drawn ones are ordinary variance charts sized to the row pitch and positioned against the rows they belong to.

| | |
|---|---|
| Cells used | `A1:H35` |
| Print area | `'T03A'!$A$9:$H$35` |
| Formula cells | 114, in 20 shapes |
| Typed number cells | 42 |

**Typed values** - 42 cells in `B15:D34`. Everything else on the sheet is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in this range, so go by the shading rather than by the count.

**The formulas behind everything else.** Cells that say the same thing about their own position are one row here; the example is the first of them, in ordinary A1.

| Cells | Named | Example | Formula |
|---|---|---|---|
| `E15:E35` | + Licences | `E15` | `=D15-B15` |
| `F15:F35` | ΔPY | `F15` | `=IF(B15=0,NA(),(D15-B15)/B15*100)` |
| `G15:G35` | + Licences | `G15` | `=D15-C15` |
| `H15:H35` | ΔPL | `H15` | `=IF(C15=0,NA(),(D15-C15)/C15*100)` |
| `B19:D19` | - | `B19` | `=SUM(B15:B18)` |
| `B26:D26` | - Operating expenses | `B26` | `=ABS(-SUM(B21:B25))` |
| `B27:D27` | - | `B27` | `=SUM(B15:B18)+B20-SUM(B21:B25)` |
| `B28:D28` | Gross margin | `B28` | `=IF(B19=0,NA(),B27/B19*100)` |
| `B31:D31` | - | `B31` | `=SUM(B15:B18)+B20-SUM(B21:B25)+SUM(B29:B30)` |
| `B33:D33` | - | `B33` | `=SUM(B15:B18)+B20-SUM(B21:B25)+SUM(B29:B30)-B32` |
| `B35:D35` | - | `B35` | `=SUM(B15:B18)+B20-SUM(B21:B25)+SUM(B29:B30)-B32-B34` |
| `B7` | Subject | `B7` | `=B2&IF(B3="",""," in "&B3)` |
| `A19` | + Other revenue | `A19` | `= Sales revenue` |
| `A27` | - Operating expenses | `A27` | `= Operating result` |
| `A31` | + Financial income, net | `A31` | `= Result before tax` |
| `A33` | - Income tax | `A33` | `= Result after tax` |
| `A35` | - Profit to other investors | `A35` | `= Group result` |

3 unnamed one-off formulas: `A9`, `A10`, `A11`.

### T04A

**Table.** A printed grid, with some columns *drawn* instead of printed - the drawn ones are ordinary variance charts sized to the row pitch and positioned against the rows they belong to.

| | |
|---|---|
| Cells used | `A1:P35` |
| Print area | `'T04A'!$A$9:$F$36` |
| Formula cells | 233, in 31 shapes |
| Typed number cells | 28 |

**Typed values** - 28 cells in `B15:C34`. Everything else on the sheet is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in this range, so go by the shading rather than by the count.

**The formulas behind everything else.** Cells that say the same thing about their own position are one row here; the example is the first of them, in ordinary A1.

| Cells | Named | Example | Formula |
|---|---|---|---|
| `I15:I35` | ΔPL scaled | `I15` | `=IF(ISBLANK(D15),NA(),D15/$G$15)` |
| `J15:J35` | ΔPL text | `J15` | `=IF(ISBLANK(D15),"",TEXT(D15,"[>0.05]+0.0;[<-0.05]-0.0;0.0"))` |
| `M15:M35` | ΔPL% scaled | `M15` | `=IF(ISBLANK(E15),NA(),E15/$H$15)` |
| `N15:N35` | ΔPL% text | `N15` | `=IF(ISBLANK(E15),"",TEXT(E15,"[>0.05]+0.0""%"";[<-0.05]-0.0""%"";0.0""%"""))` |
| 20 cells in `D15:D35` | ΔPL | `D15` | `=C15-B15` |
| 20 cells in `E15:E35` | ΔPL% | `E15` | `=IF(B15=0,NA(),(C15-B15)/B15*100)` |
| 13 cells in `K15:K35` | ΔPL good | `K15` | `=IF(ISBLANK(D15),NA(),IF(D15>0,D15/$G$15,NA()))` |
| 13 cells in `L15:L35` | ΔPL bad | `L15` | `=IF(ISBLANK(D15),NA(),IF(D15<0,D15/$G$15,NA()))` |
| 13 cells in `O15:O35` | ΔPL% good | `O15` | `=IF(ISBLANK(E15),NA(),IF(E15>0,E15/$H$15,NA()))` |
| 13 cells in `P15:P35` | ΔPL% bad | `P15` | `=IF(ISBLANK(E15),NA(),IF(E15<0,E15/$H$15,NA()))` |
| 8 cells in `K21:K34` | Purchases | `K21` | `=IF(ISBLANK(D21),NA(),IF(D21<0,D21/$G$15,NA()))` |
| 8 cells in `L21:L34` | Purchases | `L21` | `=IF(ISBLANK(D21),NA(),IF(D21>0,D21/$G$15,NA()))` |
| 8 cells in `O21:O34` | Purchases | `O21` | `=IF(ISBLANK(E21),NA(),IF(E21<0,E21/$H$15,NA()))` |
| 8 cells in `P21:P34` | Purchases | `P21` | `=IF(ISBLANK(E21),NA(),IF(E21>0,E21/$H$15,NA()))` |
| `G15:H15` | span unit, span rel | `G15` | `=MAX(AGGREGATE(4,6,D15:D35),-AGGREGATE(5,6,D15:D35),0.000000001)` |
| `B19:C19` | - | `B19` | `=SUM(B15:B18)` |
| `B26:C26` | - Operating expenses | `B26` | `=ABS(-SUM(B21:B25))` |
| `B27:C27` | - | `B27` | `=SUM(B15:B18)+B20-SUM(B21:B25)` |
| `B28:C28` | Gross margin | `B28` | `=IF(B19=0,NA(),B27/B19*100)` |
| `B31:C31` | - | `B31` | `=SUM(B15:B18)+B20-SUM(B21:B25)+SUM(B29:B30)` |
| `B33:C33` | - | `B33` | `=SUM(B15:B18)+B20-SUM(B21:B25)+SUM(B29:B30)-B32` |
| `B35:C35` | - | `B35` | `=SUM(B15:B18)+B20-SUM(B21:B25)+SUM(B29:B30)-B32-B34` |
| `B7` | Subject | `B7` | `=B2&IF(B3="",""," in "&B3)` |
| `A19` | + Other revenue | `A19` | `= Sales revenue` |
| `A27` | - Operating expenses | `A27` | `= Operating result` |
| `A31` | + Financial income, net | `A31` | `= Result before tax` |
| `A33` | - Income tax | `A33` | `= Result after tax` |
| `A35` | - Profit to other investors | `A35` | `= Group result` |

3 unnamed one-off formulas: `A9`, `A10`, `A11`.

**Chart objects**, in points from the top left of the sheet. The positions are what make separate charts read as one figure, so they are declared rather than dragged.

| Name | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|
| `panel_dpl` | 308.1 | 205.1 | 244.4 | 318.9 | 3 | -1.0288 .. 1.5167 |
| `panel_dplp` | 548.6 | 205.1 | 244.4 | 318.9 | 3 | -1.1775 .. 1.593 |

An axis range that is not round - `0.9` rather than `900` - is the scale block at work: the bounds were divided by the span cell, exactly as the values were. Multiply an axis bound by its span and you get back the range the geometry was measured in.

---

*Generated by `ibcs_doc.py` from the workbook itself.*
