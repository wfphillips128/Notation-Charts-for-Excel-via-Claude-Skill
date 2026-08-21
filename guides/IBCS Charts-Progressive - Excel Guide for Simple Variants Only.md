# Building `IBCS Charts-Progressive - Simple Variants Only.xlsx` by hand in Excel

This is the same workbook, done manually - no macros, no add-ins, no code. Work through Part 1 once and you have the sheet every template in the library sits on; Part 2 builds the charts on top of it.

**Every range, formula, colour, gap width and axis bound below was read back out of `IBCS Charts-Progressive - Simple Variants Only.xlsx` after it was written.** Where this says a formula sits in `C03A!E13`, it does; where it says a fill is `#404040`, that is the value in the file. The prose was written once. The numbers are not typed twice, which is the only way a document like this stays true to the thing it describes.

This is the guide to the **simple** workbook - each template reduced to its base tiers. The construction is identical to the full workbook's; there is simply less of it on each sheet.

## What you need

| | |
|---|---|
| Excel | 2016 or later, Windows or Mac. Nothing here needs 365 |
| Time | about an hour for your first tier stack, twenty minutes after |
| Skills | writing a formula, inserting a chart, and `Ctrl+1` |

`Ctrl+1` opens the Format pane for whatever is selected, and it is the single most useful key in this whole exercise. `Cmd+1` on a Mac.

Two habits will save you most of the pain:

- **Use the Current Selection dropdown.** Chart Format tab, far left. Series in these charts are deliberately thin, invisible, or stacked under each other, and hunting for them with the mouse is miserable. Pick them from the list by name instead.
- **Name every chart as you create it.** Click the chart, then type the name into the Name Box (left of the formula bar) and press Enter. The names used below are the ones in the workbook, and by the time you have three overlapping charts on a sheet you will want them.

## The one idea

**A chart in this notation is several charts, stacked, sharing one category axis.**

A measure and a variance are different units and different scales, so they cannot share a value axis - but they describe the same twelve months, so they must line up. Excel will not do that inside one chart object. So each tier is its own chart, sized and positioned in points so that category 1 of the top tier sits directly above category 1 of the bottom one.

```
   Delta PY%    |  pins, own scale, own chart object
   Delta PY     |  columns, own scale, own chart object
   Contribution |  columns, own scale, own chart object
                 Jan Feb Mar Apr ...   <- one shared category axis
```

Everything else in this guide is in service of that, or of the second idea: **the reader edits cells, never the chart.**

## The notation you are reproducing

IBCS is a notation, not a style guide: the same meaning must always take the same visual form. Four fills carry most of it, and these are the values in this workbook:

| Scenario | Means | Fill | Border |
|---|---|---|---|
| AC | actual - it happened | solid `#404040` | none |
| PY | prior year - it happened, earlier | solid `#A6A6A6` | none |
| PL / BU | plan or budget - fictitious | **no fill** | `#404040` |
| FC | forecast - expected | **Wide upward diagonal**, `#404040` on `#F2F2F2` | `#404040` |

| Variance | Means | Fill |
|---|---|---|
| good | favourable impact | `#8CB400` |
| bad | adverse impact | `#FF0000` |
| neutral | no direction to it | `#0064FF` |

> **Colour is by impact, not by sign.** A cost 40 over plan is a positive number and a bad outcome, so it is red. If you colour on `>0` you will get every cost overrun green, and the chart will be confidently wrong.

> **IBCS prescribes no colour codes.** Rule UN 4.1 says so outright. These hex values are this project's house palette, not a requirement of the standard.

## What is in this workbook

| Sheet | What it is | Charts |
|---|---|---|
| [Sources](#sources) | a tier stack | 0 |
| [C01A](#c01a) | a structure chart - net premiums written | 1 |
| [C02A](#c02a) | a structure chart - net premiums earned | 1 |
| [C03A](#c03a) | a tier stack - net premiums written | 1 |
| [C04A](#c04a) | a tier stack - underwriting result | 1 |
| [C05X](#c05x) | a tier stack - underwriting result | 2 |
| [C06F](#c06f) | a tier stack - net premiums written | 2 |
| [C12A](#c12a) | a tier stack - profit and loss statement | 1 |

---

# Part 1 - The skeleton every sheet shares

Seven steps. They are the same seven whichever chart you are going to build on top, so they are here once rather than seven times. `C03A` is used for the addresses; every other sheet is the same shape with different columns.

## Step 1 - Six typed lines at the top, and nothing else typed

Rows 1 to 6 of column A are labels, and column B is the only text anybody ever types:

| Row | A | B |
|---|---|---|
| 1 | Entity | The Progressive Corporation |
| 2 | Measure | Net premiums written |
| 3 | Unit | $m |
| 4 | Period | 2026 |
| 5 | Reference | PY |
| 6 | Message | Net premiums written are running +2.5% to +9.6% above prior year through Jul; on that pace ... |

Row 7 is the **subject line**, and it is a formula:

```
=B2&" in "&B3
```

Read from `C03A!B7`.

Then the tier captions, one per drawn tier that needs one:

*(nothing in this workbook does this)*

This is the part people skip, and it is the part that makes the sheet worth keeping. **Nothing a chart says may be typed into the chart.** Every text box on the finished chart is *linked* to one of these cells, so retyping `kEUR` as `kGBP` in `B3` changes the subject line and every caption at once. Type it into the chart and it is a lie the first time the data changes.

To link a text box: draw it, then with the box selected click in the **formula bar**, type `=` and click the cell. Not into the box - into the formula bar. It is the one place in Excel where that is the whole technique.

## Step 2 - Two zones, and nothing straddles them

The sheet is a **data zone** on the left of declared width, and the **charts** to its right. The print area is the chart zone alone - the data is the input, not the deliverable.

| Sheet | Cells used | Print area |
|---|---|---|
| Sources | `A1:D20` | `-` |
| C01A | `A1:F53` | `G1:M68` |
| C02A | `A1:V47` | `W1:AI10` |
| C03A | `A1:AC24` | `AD1:AO21` |
| C04A | `A1:Y29` | `Z1:AI30` |
| C05X | `A1:AU30` | `AV1:BI34` |
| C06F | `A1:AM30` | `AN1:BC34` |
| C12A | `A1:AU30` | `AV1:BD33` |

> **Never AutoFit a column a chart is positioned against.** The charts are placed in points, measured from where the data zone ends. One long customer name and AutoFit moves that boundary, and the charts end up drawn on top of the source data. Set column widths by hand: **Home -> Format -> Column Width**.

## Step 3 - The data block

On `C03A` the headers are on row **10** (`A10:AC10`) and the data runs rows **11 to 22**.

| Column | Header | Typed or derived |
|---|---|---|
| A | Period | **typed** (text) |
| B | Scenario | **typed** (text) |
| C | PY | **typed** |
| D | Measure | **typed** |
| E | ΔPY | formula |
| F | ΔPY% | formula |
| G | Measure AC | formula |
| H | Measure FC | formula |
| I | ΔPY up | formula |
| J | ΔPY down | formula |
| K | ΔPY% up | formula |
| L | ΔPY% down | formula |
| M | up length | formula |
| N | down length | formula |
| ... | 15 more, all formulas | |

Two columns earn special mention because they are notation rather than data:

- **A scenario column.** One of `AC`, `PY`, `PL`, `BU`, `FC` per row. This is what later decides which fill a bar gets, so it is a *value* in a cell and not a colour somebody applied. That is the whole reason a forecast can turn into an actual by typing over one cell.
- **A sign or kind column**, on the waterfall templates: `+1` / `-1`, or `total` / `state` / `variance`. The cascade formulas read it.

Shade the typed cells so a reader can see what is theirs to edit, and leave everything else unshaded - including the subtotals, which is the point.

## Step 4 - Split every drawn series by what it means

This is the step that surprises people, and everything downstream depends on it.

**Excel gives one colour to a series.** So a variance that is green when favourable and red when adverse cannot be one series - it has to be two, split in the *cells*, each holding `NA()` where the other owns the point:

```
=IF($E11>0,$E11,NA())
```

Read from `C03A!I11`.

Do the same wherever a fill changes for any reason: actual against forecast, plan against budget, an increase against a decrease.

> ### `NA()`, never `""`
>
> An empty string is **not** an empty cell. Excel plots it as a **zero**, so every gap becomes a bar of nothing sitting on the axis. `NA()` is a genuine gap. Also set **Select Data -> Hidden and Empty Cells -> Show empty cells as: Gaps**.
>
> This is the single most common way to end up with a chart that looks broken for no visible reason.

And once more, because it is the difference between a correct chart and a confident lie: **split by impact, not by sign.** Within one variance column of a P&L the favourable values sit on both sides of the axis, because a cost line going up is adverse and a cost line going down is favourable.

## Step 5 - The scale block, so the charts follow *your* numbers

Skip this one and everything still works - until somebody pastes their own figures in, and every bar clips or shrinks to nothing.

**Excel will not bind an axis bound to a formula.** An axis maximum is a number you type into a box; it cannot be `=MAX(...)`. So a chart built for one set of figures is built for *those* figures.

The way round it is to turn the problem over. Instead of scaling the axis to the data, **scale the data to a fixed axis**. Three parts:

**1. A span cell** - one per group of tiers that share a unit - holding the largest magnitude anywhere in that group:

```
=MAX(AGGREGATE(4,6,C11:C22,D11:D22,E11:E22,I11:I22,J11:J22),-AGGREGATE(5,6,C11:C22,D11:D22,E11:E22,I11:I22,J11:J22),0.000000001)
```

Read from `C03A!O11`. `AGGREGATE(4,6,...)` is MAX ignoring errors, which matters because half of these columns are deliberately `NA()`.

**2. A scaled copy of every column a chart reads**, which is what you actually plot:

```
=C11/$O$11
```

Read from `C03A!Q11`.

**3. Axis bounds divided by the same span.** If your measure tier was designed to run 0 to 230, you type `0` and `230/span` - or rather, you work out that number once and type the result.

Dividing both the values and the bounds by one number is a visual no-op at today's figures and follows the data at any others. And because zero divided by anything is zero, the zero line, the reference rules and every caption positioned from plot geometry stay exactly where they were.

Hide the scaled columns - **right-click the column headers -> Hide**. A hidden column measures zero width, so you can add all of this to a sheet without moving a single chart on it.

> **Where this cannot be used.** Every chart in this library hides its value axis, which is why dividing by a span is invisible. The two XY sheets are the exception - they show gridlines and tick labels, and dividing those by a span would print `0.25` where the data says `29.16%`. Those two get their axis fitted to the data once, at build time, and do not follow an edit. If you are building a chart whose axis a reader actually reads, do the same.

## Step 6 - A text column for every label

A data label linked to a cell shows the **cell's** value and ignores the label's own number format. Link a label to a raw variance and you get `31.61057692` printed beside a chart that says `+32`.

So the formatting happens in a cell:

```
=IF(ISNA(D11),"",TEXT(D11,"0"))
```

Read from `C03A!S11`.

One such column per labelled series, blank where that series has no point. Then in step 2.9 you point the labels at these columns with **Value From Cells**.

This is also why you cannot skip it: the label range applies to the whole series, so a shared column would print the other series' numbers on your bars.

## Step 7 - Page setup, before you forget

1. Select the range the charts cover, then **Page Layout -> Print Area -> Set Print Area**.
2. **Page Layout -> Orientation**, then **Scale to Fit: Width 1 page, Height 1 page**.
3. **View -> uncheck Gridlines**, so the worksheet grid does not print behind the charts.
4. Check nothing else is inside the print area. Cell contents under a chart print straight through it.

---

# Part 2 - Worked builds

One per family present in this workbook. The tier stack is done in full; the rest are written as what differs from it.

## Building a tier stack, on `C03A`

The full walkthrough. Every other family in Part 2 is written as *what differs from this*, so read it even if the chart you want is a different shape.

`C03A` is net premiums written over 12 categories, in 1 tier.

### 2.1 The columns, and what each is for

Headers on row **10**, data rows **11-22**. Typed cells are `C11:D22` - everything else is a formula.

| Column | Header | Formula in row 11 |
|---|---|---|
| A | Period | **typed** - `Jan`, and down |
| B | Scenario | **typed** - `AC`, and down |
| C | PY | **typed** - `6481`, and down |
| D | Measure | **typed** - `6735`, and down |
| E | ΔPY | `=D11-C11` |
| F | ΔPY% | `=IF(C11<=0,NA(),(D11-C11)/C11*100)` |
| G | Measure AC | `=IF($B11="AC",$D11,NA())` |
| H | Measure FC | `=IF($B11="AC",NA(),$D11)` |
| I | ΔPY up | `=IF($E11>0,$E11,NA())` |
| J | ΔPY down | `=IF($E11<0,$E11,NA())` |
| K | ΔPY% up | `=IF($F11>0,$F11,NA())` |
| L | ΔPY% down | `=IF($F11<0,$F11,NA())` |
| M | up length | `=IF($F11>0,$F11,0)` |
| N | down length | `=IF($F11<0,-$F11,0)` |
| O | span unit | `=MAX(AGGREGATE(4,6,C11:C22,D11:D22,E11:E22,I11:I22,J11:J22),-AGGREGATE(5,6,C11:C22,D11:D22,E11:E22,I11:I22,J11:J22),0.000000001)` |
| P | span rel | `=MAX(AGGREGATE(4,6,K11:K22,L11:L22,M11:M22,N11:N22),-AGGREGATE(5,6,K11:K22,L11:L22,M11:M22,N11:N22),0.000000001)` |
| Q | ref scaled | `=C11/$O$11` |
| R | measure scaled | `=D11/$O$11` |
| S | measure text | `=IF(ISNA(D11),"",TEXT(D11,"0"))` |
| T | var_abs scaled | `=E11/$O$11` |
| U | var_abs text | `=IF(ISNA(E11),"",TEXT(E11,"+0;-0"))` |
| V | abs_up scaled | `=I11/$O$11` |
| W | abs_dn scaled | `=J11/$O$11` |
| X | rel_up scaled | `=K11/$P$11` |
| Y | rel_up text | `=IF(ISNA(K11),"",TEXT(K11,"+0.0;-0.0"))` |
| Z | rel_dn scaled | `=L11/$P$11` |
| AA | rel_dn text | `=IF(ISNA(L11),"",TEXT(L11,"+0.0;-0.0"))` |
| AB | rel_up_len scaled | `=M11/$P$11` |
| AC | rel_dn_len scaled | `=N11/$P$11` |

Enter each formula in row 11 and fill down to row 22.

### 2.2 Check the numbers before you go near a chart

If these are right the chart cannot go far wrong; if they are wrong no amount of formatting will save it.

| Row | Period | Scenario | PY | Measure | ΔPY | ΔPY% | Measure AC | Measure FC |
|---|---|---|---|---|---|---|---|---|
| 11 | Jan | AC | 6481 | 6735 | `=D11-C11` | `=IF(C11<=0,NA(),(D11-C11)/C11*100)` | `=IF($B11="AC",$D11,NA())` | `=IF($B11="AC",NA(),$D11)` |
| 12 | Feb | AC | 6684 | 6995 | `=D12-C12` | `=IF(C12<=0,NA(),(D12-C12)/C12*100)` | `=IF($B12="AC",$D12,NA())` | `=IF($B12="AC",NA(),$D12)` |
| 13 | Mar | AC | 9041 | 9911 | `=D13-C13` | `=IF(C13<=0,NA(),(D13-C13)/C13*100)` | `=IF($B13="AC",$D13,NA())` | `=IF($B13="AC",NA(),$D13)` |
| 14 | Apr | AC | 6837 | 7278 | `=D14-C14` | `=IF(C14<=0,NA(),(D14-C14)/C14*100)` | `=IF($B14="AC",$D14,NA())` | `=IF($B14="AC",NA(),$D14)` |
| ... |  |  |  |  |  |  |  |  |

### 2.3 Insert the bottom tier

The measure tier is the one everything else is aligned to, so it goes first.

1. Select the category labels and the scaled measure columns together - hold **Ctrl** for the second block, and include the header row so the series get their names.
2. **Insert -> Clustered Column.**
3. Name it `tier_measure` in the Name Box.
4. `Ctrl+1` on the chart area -> **Size**: width **520 pt**, height **244 pt**. Under **Properties**, set **Don't move or size with cells** - otherwise inserting a row later moves your chart.

Then the settings that are not defaults:

| Where | Setting | Value |
|---|---|---|
| any series -> Series Options | Gap Width | **23%** |
| same pane | Series Overlap | **76%** |
| value axis -> Axis Options | Minimum / Maximum | **0** / **1.09** |
| value axis -> Labels | Label Position | **None** |
| value axis -> Line | | **No line**, no tick marks |
| chart area -> Border | | **No line** |
| Chart Elements (+) | Gridlines, Legend, Title | **all off** |

> The axis bounds are not round numbers because of the scale block in Part 1 step 5: they are the range the geometry was designed in, divided by the span cell. Multiply them back by the span and you get the round numbers you started from.

### 2.4 The fills - this is the notation

Select each series from **Chart Format -> Current Selection**, then `Ctrl+1` -> Fill & Line:

| Series | Fill | Line |
|---|---|---|
| `PY` | solid #A6A6A6 | no line |
| `AC` | solid #404040 | no line |

**The forecast is per point, not per series.** On `AC` the series carries the actual fill and the forecast months are overridden individually:

| Points | Fill |
|---|---|
| 1-7 | solid #404040 |
| 8-12 | Wide upward diagonal, #404040 on #F2F2F2 |

To do it: click the series once to select all of it, then click **again** on the single bar you want - that selects the point - then `Ctrl+1` -> Fill -> **Pattern fill**, and pick the pattern from the gallery with the foreground and background above.

> Yes, this is manual, and yes it is the one thing on the sheet that does not follow the data. If the forecast starts a month earlier next quarter you must re-apply it. The alternative is a separate series for the forecast points, driven by the scenario column - more columns, no clicking. Both are used in this workbook; look at which sheets have an `AC` and an `FC` series and which have one series with overridden points.

### 2.7 Line the tiers up - the step that makes it one figure

Three charts that nearly line up look like a mistake. They have to be exact, and you get there by typing numbers, not by dragging.

For each chart: `Ctrl+1` -> **Size & Properties**, and set Height, Width, and under Position the Horizontal and Vertical offsets:

| Chart | Left | Top | Width | Height |
|---|---|---|---|---|
| `tier_measure` | 683 | 86 | 520 | 244 |

That is not enough on its own. Two charts of the same width can still have plot areas of different widths, because Excel sizes the plot area around whatever labels it has to fit. So set the **plot area** too: click inside the plot (not the chart), `Ctrl+1`, and set its size and position:

| Chart | Plot left | Plot top | Plot width | Plot height |
|---|---|---|---|---|
| `tier_measure` | 46 | 6 | 452 | 210 |

> **Set the plot area size before the position, and check it after.** Excel treats a plot-area assignment as a request: it will quietly adjust what you asked for to fit the labels, and it applies width and position in an order that means setting one can undo the other. Set them, then look at the numbers again.

Only the bottom tier keeps its category axis labels. On the others, select the category axis -> Labels -> **Label Position: None**. The labels are shared, so they are printed once.

### 2.8 Captions and the title block

Draw a text box for each of the four title lines and each tier caption, and link every one of them to its cell as in Part 1 step 1 - select the box, click in the formula bar, type `=`, click the cell.

Set **Word Wrap off** on the captions, and turn off **Resize shape to fit text**. A caption that rewraps when the unit changes from `kEUR` to `kUSD` will push itself over the chart.

### 2.9 The data labels

For each series that carries labels:

1. Select the series -> **Chart Elements (+) -> Data Labels**.
2. `Ctrl+1` -> Label Options -> tick **Value From Cells**, and select that series' text column from Part 1 step 6.
3. **Untick Value**, and untick everything else. Only Value From Cells stays on.
4. Position: **Outside End** for columns that grow from an axis, **Inside End** where the label would otherwise leave the plot.

| Series | Labelled |
|---|---|
| `tier_measure` / `AC` | yes |

### 2.10 Check it before you send it

Numbers first, appearance second.

- [ ] **Every scenario has the right fill.** Solid dark for actual, solid light for prior year, outlined for plan, hatched for forecast. One wrong fill and the chart says something untrue.
- [ ] **Variance colour follows impact, not sign.** Check a cost line that went up: it must be red.
- [ ] **The tiers line up.** Put a ruler on the screen - category 1 of the top tier over category 1 of the bottom.
- [ ] **The variance axis carries its reference.** A reader must be able to tell what the variance is *from* without a legend.
- [ ] **No stray zero labels** where a series has no point - that is the `""`-instead-of-`NA()` symptom.
- [ ] **Every label is linked**, not typed. Change a number and watch them all move.
- [ ] **Retype the unit in `B3`.** The subject line and every caption must follow. If one does not, it was typed into the chart.
- [ ] **Nothing prints under the charts.** Print Preview, one page.

### 2.11 What goes wrong, and why

| Symptom | Cause | Fix |
|---|---|---|
| A bar of nothing sitting on the axis | a formula returns `""` where it means `NA()` | step 4 |
| Every cost overrun is green | the split is on sign, not impact | step 4 |
| Bars clip the moment new data is pasted | axis bounds typed as data values with no scale block | step 5 |
| A label prints `31.61057692` | linked to the raw cell, not a `TEXT()` column | step 6 |
| Labels stop updating | typed into the chart instead of linked | step 1 |
| Tiers nearly line up | plot area sizes not set, only chart sizes | step 2.7 |
| Charts drawn over the data | AutoFit moved the zone boundary | step 2 |
| A pin has a T on the end | error-bar End Style is not No Cap | step 2.6 |
| Stray lines between pin heads | a point's Line was styled instead of its Marker Border | step 2.6 |
| An error bar does nothing | its range contains an `NA()` or a blank | step 2.6 |
| The whole chart re-scales when a row is inserted | chart Properties left on Move and size with cells | step 2.3 |

---

## Building a structure chart, on `C01A`

Part 1 applies unchanged. What differs:

- **One data block per panel**, side by side, and every panel on **one shared scale** - which is the entire point of the template. Work out the tallest column across all panels, round it up, and give every panel that maximum. Do not let Excel scale each one.
- **A stacked category is not a stacked scenario.** Where the bands are business areas, every band is an actual, so the fill cannot carry the scenario any more. The *column* carries it - a plan column is outlined - and the bands take a light-to-dark ramp.
- **Insert as a Stacked Column (or Bar)**, one series per band, bottom band first.
- **Small bands get no label.** Below about 3%% of the axis a number does not fit, and a label that overlaps its neighbour is worse than an absent one. That test belongs in the label formula, not in your judgement:

```
=IF(ISNA(D11),"",TEXT(D11,"0"))
```

Read from `C03A!S11`.


**The charts on this sheet, as built:**

| Chart | Type | Left | Top | Width | Height | Series |
|---|---|---|---|---|---|---|
| `panel_segment` | chart type -4111 | 407.5 | 81 | 300 | 430 | 6 |

---

# Part 3 - Sheet by sheet

What is on each sheet, for rebuilding one template rather than learning the method.

### Sources

> AC

**A tier stack.**

| | |
|---|---|
| Cells used | `A1:D20` |
| Print area | `-` |
| Formula cells | 0, in 0 shapes |
| Typed cells | 0 |

**Every formula on the sheet.** Cells that say the same thing about their own position are one row here, with the range they cover and the first of them written out.

| Cells | Named | Example | Formula |
|---|---|---|---|

### C01A

> Direct auto has grown from 18,911 to 39,631 $m since 2021, +110%, and now writes 48% of premium; Texas and Florida together account for 25% of the 2025 book

**A structure chart.**

| | |
|---|---|
| Cells used | `A1:F53` |
| Print area | `G1:M68` |
| Formula cells | 67, in 6 shapes |
| Typed cells | 25 |

**Typed values** - `B10:F14`. Everything else is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in that range, so go by the shading.

**Every formula on the sheet.** Cells that say the same thing about their own position are one row here, with the range they cover and the first of them written out.

| Cells | Named | Example | Formula |
|---|---|---|---|
| `B29:F34` | - | `B29` | `=IF(ISBLANK(B10),NA(),B10/$B$17)` |
| 25 cells in `B48:F53` | - | `B48` | `=IF(ISBLANK(B10),"",IF(B10/83180*264<(LEN(TEXT(B10,"# ##0")))*5,"",TEXT(B10,"# ##0")))` |
| `B15:F15` | Total | `B15` | `=SUM(B10:B14)` |
| `B52:F52` | - | `B52` | `=IF(ISBLANK(B14),"",IF(B14/83180*264<(LEN(TEXT(B14,"# ##0"))+LEN(TEXT(B15,"# ##0")))*5,"",TEXT(B14,"# ##0")))` |
| `B7` | Subject | `B7` | `=B2&IF(B3="",""," in "&B3)` |
| `B17` | span | `B17` | `=MAX(AGGREGATE(4,6,B15:F15),0.000000001)` |

**Charts**, in points from the top left of the sheet.

| Name | Type | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|---|
| `panel_segment` | chart type -4111 | 407.5 | 81 | 300 | 430 | 6 | 0 to 1.0001 |

| Chart | Series | Fill | Marker | Stem | Labelled |
|---|---|---|---|---|---|
| `panel_segment` | `Agency auto` | solid #404040 | None | - | yes |
| `panel_segment` | `Direct auto` | solid #595959 | None | - | yes |
| `panel_segment` | `Property` | solid #7F7F7F | None | - | yes |
| `panel_segment` | `Commercial` | solid #BFBFBF | None | - | yes |
| `panel_segment` | `Other` | solid #D9D9D9 | None | - | yes |
| `panel_segment` | `Total` | solid #000000 | None | - | yes |

### C02A

> Of 81,659 $m earned in 2025, 53,866 went to losses and 17,523 to expenses, leaving an underwriting result of 10,270 $m (12.6%); Property keeps the largest share of its premium at 24.9% and Direct auto the smallest at 9.9%

**A structure chart.**

| | |
|---|---|
| Cells used | `A1:V47` |
| Print area | `W1:AI10` |
| Formula cells | 38, in 6 shapes |
| Typed cells | 12 |

**Typed values** - `B10:E12`. Everything else is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in that range, so go by the shading.

**Every formula on the sheet.** Cells that say the same thing about their own position are one row here, with the range they cover and the first of them written out.

| Cells | Named | Example | Formula |
|---|---|---|---|
| `B27:E30` | - | `B27` | `=IF(ISBLANK(B10),NA(),B10/$B$15)` |
| 12 cells in `B44:E47` | - | `B44` | `=IF(ISBLANK(B10),"",IF(B10/38300*492.8<(LEN(TEXT(B10,"# ##0")))*5,"",TEXT(B10,"# ##0")))` |
| `B13:E13` | Total | `B13` | `=SUM(B10:B12)` |
| `B46:E46` | - | `B46` | `=IF(ISBLANK(B12),"",IF(B12/38300*492.8<(LEN(TEXT(B12,"# ##0"))+LEN(TEXT(B13,"# ##0")))*5,"",TEXT(B12,"# ##0")))` |
| `B7` | Subject | `B7` | `=B2&IF(B3="",""," in "&B3)` |
| `B15` | span | `B15` | `=MAX(AGGREGATE(4,6,$B$13:$E$13),0.000000001)` |

**Charts**, in points from the top left of the sheet.

| Name | Type | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|---|
| `panel_premium` | Stacked Bar | 1175.5 | 81 | 560 | 99 | 4 | 0 to 1.0005 |

| Chart | Series | Fill | Marker | Stem | Labelled |
|---|---|---|---|---|---|
| `panel_premium` | `Losses and LAE` | solid #404040 | None | - | yes |
| `panel_premium` | `Expenses` | solid #7F7F7F | None | - | yes |
| `panel_premium` | `Underwriting result` | solid #D9D9D9 | None | - | yes |
| `panel_premium` | `Total` | no fill | None | - | yes |

### C03A

> Net premiums written are running +2.5% to +9.6% above prior year through Jul; on that pace the year closes +4,753 $m (+5.7%) above 2025

**A tier stack.**

| | |
|---|---|
| Cells used | `A1:AC24` |
| Print area | `AD1:AO21` |
| Hidden scale columns | `O` |
| Formula cells | 279, in 23 shapes |
| Typed cells | 24 |

**Typed values** - `C11:D22`. Everything else is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in that range, so go by the shading.

**Every formula on the sheet.** Cells that say the same thing about their own position are one row here, with the range they cover and the first of them written out.

| Cells | Named | Example | Formula |
|---|---|---|---|
| `Q11:R22` | ref scaled, measure scaled | `Q11` | `=C11/$O$11` |
| `V11:W22` | abs_up scaled, abs_dn scaled | `V11` | `=I11/$O$11` |
| `AB11:AC22` | rel_up_len scaled, rel_dn_len scaled | `AB11` | `=M11/$P$11` |
| `E11:E22` | ΔPY | `E11` | `=D11-C11` |
| `F11:F22` | ΔPY% | `F11` | `=IF(C11<=0,NA(),(D11-C11)/C11*100)` |
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

**Charts**, in points from the top left of the sheet.

| Name | Type | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|---|
| `tier_measure` | Clustered Column | 683 | 86 | 520 | 244 | 2 | 0 to 1.09 |

| Chart | Series | Fill | Marker | Stem | Labelled |
|---|---|---|---|---|---|
| `tier_measure` | `PY` | solid #A6A6A6 | None | - |  |
| `tier_measure` | `AC` | points 1-7: solid #404040; points 8-12: Wide upward diagonal, #404040 on #F2F2F2 | None | - | yes |

### C04A

> 6 of 12 months beat the 96 combined ratio target in 2023; Dec by +670 $m and Mar missed it by 441 $m, a spread of 1,111 $m across the year

**A tier stack.**

| | |
|---|---|
| Cells used | `A1:Y29` |
| Print area | `Z1:AI30` |
| Hidden scale columns | `M` |
| Formula cells | 231, in 19 shapes |
| Typed cells | 24 |

**Typed values** - `C11:D22`. Everything else is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in that range, so go by the shading.

**Every formula on the sheet.** Cells that say the same thing about their own position are one row here, with the range they cover and the first of them written out.

| Cells | Named | Example | Formula |
|---|---|---|---|
| 36 cells in `P11:U22` | measure scaled, abs_up scaled, abs_dn scaled | `P11` | `=C11/$M$11` |
| `X11:Y22` | rel_up scaled, rel_dn scaled | `X11` | `=I11/$N$11` |
| `E11:E22` | PL | `E11` | `=C11-D11` |
| `F11:F22` | ΔPL% | `F11` | `=IF(E11<=0,NA(),D11/E11*100)` |
| `G11:G22` | ΔPL up | `G11` | `=IF($D11>0,$D11,NA())` |
| `H11:H22` | ΔPL down | `H11` | `=IF($D11<0,$D11,NA())` |
| `I11:I22` | ΔPL% up | `I11` | `=IF($F11>0,$F11,NA())` |
| `J11:J22` | ΔPL% down | `J11` | `=IF($F11<0,$F11,NA())` |
| `K11:K22` | up length | `K11` | `=IF($F11>0,$F11,0)` |
| `L11:L22` | down length | `L11` | `=IF($F11<0,-$F11,0)` |
| `O11:O22` | ref scaled | `O11` | `=E11/$M$11` |
| `Q11:Q22` | measure text | `Q11` | `=IF(ISNA(C11),"",TEXT(C11,"0"))` |
| `R11:R22` | var_abs scaled | `R11` | `=D11/$M$11` |
| `S11:S22` | var_abs text | `S11` | `=IF(ISNA(D11),"",TEXT(D11,"+0;-0"))` |
| `V11:V22` | var_rel scaled | `V11` | `=F11/$N$11` |
| `W11:W22` | var_rel text | `W11` | `=IF(ISNA(F11),"",TEXT(F11,"+0;-0"))` |
| `B7` | Subject | `B7` | `=B2&" in "&B3` |
| `M11` | span unit | `M11` | `=MAX(AGGREGATE(4,6,E11:E29,C11:C29,D11:D29,G11:G29,H11:H29),-AGGREGATE(5,6,E11:E29,C11:C29,D11:D29,G11:G29,H11:H29),0.000000001)` |
| `N11` | span rel | `N11` | `=MAX(AGGREGATE(4,6,F11:F29,I11:I29,J11:J29),-AGGREGATE(5,6,F11:F29,I11:I29,J11:J29),0.000000001)` |

**Charts**, in points from the top left of the sheet.

| Name | Type | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|---|
| `tier_measure` | Clustered Bar | 631.5 | 86 | 430 | 400 | 2 | -0.35 to 1.826 |

| Chart | Series | Fill | Marker | Stem | Labelled |
|---|---|---|---|---|---|
| `tier_measure` | `PL` | solid #FFFFFF | None | - |  |
| `tier_measure` | `AC` | solid #404040 | None | - | yes |

### C05X

> The underwriting result is running +1,200 $m (+11.7%) ahead of 2025, with 7 of twelve months up and Sep contributing +985 on its own. Both years sit far above the 96 combined ratio target, drawn as the second opening column

**A tier stack.**

| | |
|---|---|
| Cells used | `A1:AU30` |
| Print area | `AV1:BI34` |
| Hidden scale columns | `W` |
| Formula cells | 688, in 60 shapes |
| Typed cells | 25 |

**Typed values** - 25 cells in `D12:E24`. Everything else is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in that range, so go by the shading.

**Every formula on the sheet.** Cells that say the same thing about their own position are one row here, with the range they cover and the first of them written out.

| Cells | Named | Example | Formula |
|---|---|---|---|
| 51 cells in `AC11:AH27` | bar_ac scaled, wf_base scaled, wf_up_ac scaled | `AC11` | `=J11/$W$11` |
| 40 cells in `G11:T27` | ΔPY%, PY bar, AC bar, FC bar | `G11` | `=NA()` |
| 34 cells in `AE11:AJ27` | bar_fc scaled, wf_dn_ac scaled | `AE11` | `=K11/$W$11` |
| `AT11:AU27` | rel_up_len scaled, rel_dn_len scaled | `AT11` | `=U11/$X$11` |
| 33 cells in `F11:V27` | PY, level, base, up AC | `F11` | `=0` |
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
| `L26:L27` | vs PY | `L26` | `=$D$25` |
| `O26:O27` | vs PY | `O26` | `=IF(E26>0,E26,0)` |

*20 further named shapes are not listed - this sheet has 60 in all.*

**Charts**, in points from the top left of the sheet.

| Name | Type | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|---|
| `tier_wf` | Stacked Column | 1038.5 | 86 | 620 | 110 | 5 | 0.7943 to 1.101 |
| `tier_measure` | Clustered Column | 1038.5 | 206 | 620 | 330 | 4 | -0.0024 to 1.02 |

| Chart | Series | Fill | Marker | Stem | Labelled |
|---|---|---|---|---|---|
| `tier_wf` | `base` | no fill | None | - |  |
| `tier_wf` | `up AC` | solid #8CB400 | None | - | yes |
| `tier_wf` | `down AC` | solid #FF0000 | None | - | yes |
| `tier_wf` | `up FC` | Wide upward diagonal, #8CB400 on #FFFFFF | None | - | yes |
| `tier_wf` | `down FC` | Wide upward diagonal, #FF0000 on #FFFFFF | None | - | yes |
| `tier_measure` | `PY bar` | solid #A6A6A6 | None | - | yes |
| `tier_measure` | `PY bar` | solid #FFFFFF | None | - | yes |
| `tier_measure` | `AC bar` | solid #404040 | None | - | yes |
| `tier_measure` | `FC bar` | Wide upward diagonal, #404040 on #F2F2F2 | None | - | yes |

### C06F

> Premium grew +8,750 $m (+11.8%) and every one of the ten largest states grew. But 56% of the increase came from outside the top ten - +4,885 $m against +3,865 from all ten named states put together, Texas included

**A tier stack.**

| | |
|---|---|
| Cells used | `A1:AM30` |
| Print area | `AN1:BC34` |
| Hidden scale columns | `T` |
| Formula cells | 452, in 41 shapes |
| Typed cells | 24 |

**Typed values** - 24 cells in `D12:E24`. Everything else is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in that range, so go by the shading.

**Every formula on the sheet.** Cells that say the same thing about their own position are one row here, with the range they cover and the first of them written out.

| Cells | Named | Example | Formula |
|---|---|---|---|
| 55 cells in `F11:Q24` | PY, AC bar, PY short, PL bar | `F11` | `=0` |
| 42 cells in `AA11:AF24` | bar_py scaled, wf_base scaled, wf_up scaled | `AA11` | `=K11/$T$11` |
| `X11:Y24` | bar_short scaled, bar_pl scaled | `X11` | `=I11/$T$11` |
| 28 cells in `AC11:AH24` | bar_ac scaled, wf_dn scaled | `AC11` | `=L11/$T$11` |
| `AL11:AM24` | rel_up scaled, rel_dn scaled | `AL11` | `=R11/$U$11` |
| `V11:V24` | bar_state scaled | `V11` | `=H11/$T$11` |
| `W11:W24` | bar_state text | `W11` | `=IF(ISNA(H11),"",TEXT(H11,"#,##0"))` |
| `Z11:Z24` | bar_pl text | `Z11` | `=IF(ISNA(J11),"",TEXT(J11,"#,##0"))` |
| `AB11:AB24` | bar_py text | `AB11` | `=IF(ISNA(K11),"",TEXT(K11,"#,##0"))` |
| `AD11:AD24` | bar_ac text | `AD11` | `=IF(ISNA(L11),"",TEXT(L11,"#,##0"))` |
| `AG11:AG24` | wf_up text | `AG11` | `=IF(ISNA(P11),"",TEXT(P11,"+#,##0"))` |
| `AI11:AI24` | wf_dn text | `AI11` | `=IF(ISNA(Q11),"",TEXT(Q11,"""-""#,##0"))` |
| `AJ11:AJ24` | var_rel scaled | `AJ11` | `=E11/$U$11` |
| `AK11:AK24` | var_rel text | `AK11` | `=IF(ISNA(E11),"",TEXT(E11,"+0;-0"))` |
| 12 cells in `P12:P24` | Texas | `P12` | `=IF(G12>0,G12,0)` |
| 12 cells in `Q12:Q24` | Texas | `Q12` | `=IF(G12<0,-G12,0)` |
| `R12:R23` | Texas | `R12` | `=IF(E12>0,E12,NA())` |
| `S12:S23` | Texas | `S12` | `=IF(E12<0,E12,NA())` |
| `F12:F22` | Texas | `F12` | `=IF(E12=-100,NA(),D12/(1+E12/100))` |
| `G12:G22` | Texas | `G12` | `=D12-F12` |
| `H12:H22` | Texas | `H12` | `=D12` |
| `I12:I22` | Texas | `I12` | `=MAX(F12-D12,0)` |
| `M12:M22` | Texas | `M12` | `=N12+G12` |
| `N12:N22` | Texas | `N12` | `=M11` |
| `O12:O22` | Texas | `O12` | `=MIN(N12,M12)` |
| 4 cells in `R11:S24` | ΔPY% up, ΔPY% down | `R11` | `=NA()` |
| 2 cells in `N11:N23` | from | `N11` | `=M11` |
| `B7` | Subject | `B7` | `=B2&" in "&B3` |
| `D11` | AC | `D11` | `=SUM(F12:F22)` |
| `K11` | PY bar | `K11` | `=D11` |
| `M11` | level | `M11` | `=D11` |
| `T11` | span unit | `T11` | `=MAX(AGGREGATE(4,6,H11:H30,I11:I30,J11:J30,K11:K30,L11:L30,O11:O30,P11:P30,Q11:Q30),-AGGREGATE(5,6,H11:H30,I11:I30,J11:J30,K11:K30,L11:L30,O11:O30,P11:P30,Q11:Q30),0.000000001)` |
| `U11` | span rel | `U11` | `=MAX(AGGREGATE(4,6,E11:E30,R11:R30,S11:S30),-AGGREGATE(5,6,E11:E30,R11:R30,S11:S30),0.000000001)` |
| `D23` | AC | `D23` | `=SUM(D12:D22)` |
| `G23` | AC | `G23` | `=D23-$D$11` |
| `L23` | AC | `L23` | `=D23` |
| `M23` | AC | `M23` | `=M22` |
| `G24` | ΔPY | `G24` | `=$D$23-$D$11` |
| `M24` | ΔPY | `M24` | `=$D$23` |
| `N24` | ΔPY | `N24` | `=$D$11` |

*1 further named shapes are not listed - this sheet has 41 in all.*

**Charts**, in points from the top left of the sheet.

| Name | Type | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|---|
| `tier_measure` | Stacked Bar | 970.5 | 86 | 560 | 430 | 5 | 0 to 1.02 |
| `tier_wf` | Stacked Bar | 1540.5 | 86 | 140 | 430 | 3 | 0.8083 to 1.0865 |

| Chart | Series | Fill | Marker | Stem | Labelled |
|---|---|---|---|---|---|
| `tier_measure` | `AC bar` | solid #404040 | None | - | yes |
| `tier_measure` | `PY short` | solid #A6A6A6 | None | - |  |
| `tier_measure` | `PL bar` | solid #FFFFFF | None | - | yes |
| `tier_measure` | `PY bar` | solid #A6A6A6 | None | - | yes |
| `tier_measure` | `AC total` | solid #404040 | None | - | yes |
| `tier_wf` | `base` | no fill | None | - |  |
| `tier_wf` | `step up` | solid #8CB400 | None | - | yes |
| `tier_wf` | `step down` | solid #FF0000 | None | - | yes |

### C12A

> Net income rose from 722 to 3,903 $m, +441%, as an underwriting result that had nearly vanished in 2022 recovered; 3 lines move too far to be drawn to scale and carry the overflow marks

**A tier stack.**

| | |
|---|---|
| Cells used | `A1:AU30` |
| Print area | `AV1:BD33` |
| Hidden scale columns | `AP`, `AU`, `Y` |
| Formula cells | 428, in 36 shapes |
| Typed cells | 24 |

**Typed values** - 24 cells in `B11:D20`. Everything else is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in that range, so go by the shading.

**Every formula on the sheet.** Cells that say the same thing about their own position are one row here, with the range they cover and the first of them written out.

| Cells | Named | Example | Formula |
|---|---|---|---|
| 40 cells in `Z11:AF20` | py_base scaled, py_add scaled, ac_base scaled, ac_add scaled | `Z11` | `=I11/$Y$11` |
| `AL11:AO20` | abs_pos_good scaled, abs_pos_bad scaled, abs_neg_good scaled, abs_neg_bad scaled | `AL11` | `=Q11/$Y$11` |
| `AQ11:AT20` | rel_pos_good clipped, rel_pos_bad clipped, rel_neg_good clipped, rel_neg_bad clipped | `AQ11` | `=IF(ISNA(U11),NA(),MEDIAN(-31.5,U11,238.5))` |
| 20 cells in `I11:N20` | PY base, AC base | `I11` | `=MIN(H11,G11)` |
| 20 cells in `AB11:AG20` | py_add text, ac_add text | `AB11` | `=IF(ISNA(J11),"",TEXT(J11,"#,##0"))` |
| 20 cells in `AC11:AH20` | py_sub scaled, ac_sub scaled | `AC11` | `=K11/$Y$11` |
| 20 cells in `AD11:AI20` | py_sub text, ac_sub text | `AD11` | `=IF(ISNA(K11),"",TEXT(K11,"#,##0"))` |
| 12 cells in `H12:M19` | Investment income | `H12` | `=G11` |
| `E11:E20` | ΔPY | `E11` | `=D11-C11` |
| `F11:F20` | ΔPY% | `F11` | `=IF(C11<=0,NA(),(D11-C11)/C11*100)` |
| `J11:J20` | PY adds | `J11` | `=IF(B11>0,ABS(G11-H11),0)` |
| `K11:K20` | PY subs | `K11` | `=IF(B11<0,ABS(G11-H11),0)` |
| `O11:O20` | AC adds | `O11` | `=IF(B11>0,ABS(L11-M11),0)` |
| `P11:P20` | AC subs | `P11` | `=IF(B11<0,ABS(L11-M11),0)` |
| `Q11:Q20` | ΔPY up ok | `Q11` | `=IF(AND($E11>0,$B11>0),$E11,NA())` |
| `R11:R20` | ΔPY up bad | `R11` | `=IF(AND($E11>0,$B11<0),$E11,NA())` |
| `S11:S20` | ΔPY dn ok | `S11` | `=IF(AND($E11<0,$B11<0),$E11,NA())` |
| `T11:T20` | ΔPY dn bad | `T11` | `=IF(AND($E11<0,$B11>0),$E11,NA())` |
| `U11:U20` | ΔPY% up ok | `U11` | `=IF(AND($F11>0,$B11>0),$F11,NA())` |
| `V11:V20` | ΔPY% up bad | `V11` | `=IF(AND($F11>0,$B11<0),$F11,NA())` |
| `W11:W20` | ΔPY% dn ok | `W11` | `=IF(AND($F11<0,$B11<0),$F11,NA())` |
| `X11:X20` | ΔPY% dn bad | `X11` | `=IF(AND($F11<0,$B11>0),$F11,NA())` |
| `AJ11:AJ20` | var_abs scaled | `AJ11` | `=E11/$Y$11` |
| `AK11:AK20` | var_abs text | `AK11` | `=IF(ISNA(E11),"",TEXT(E11,"+0;-0"))` |
| `AP11:AP20` | var_rel clipped | `AP11` | `=IF(ISNA(F11),NA(),MEDIAN(-31.5,F11,238.5))` |
| `AU11:AU20` | ΔPY% text | `AU11` | `=IF(ISNA(F11),"",IF(F11>265,TEXT(F11,"+0;-0")&REPT(UNICHAR(9658),MIN(3,MAX(1,INT(F11/265)))),IF(F11<-35,REPT(UNICHAR(9668),MIN(3,MAX(1,INT(-F11/35))))&TEXT(F11,"+0;-0"),TEXT(F11,"+0;-0"))))` |
| 8 cells in `H11:M20` | PY from, AC from | `H11` | `=0` |
| 6 cells in `G12:G19` | Investment income | `G12` | `=G11+B12*C12` |
| 6 cells in `L12:L19` | Investment income | `L12` | `=L11+B12*D12` |
| 6 cells in `G14:L20` | Total revenues | `G14` | `=G13` |
| 3 cells in `C14:C20` | Total revenues | `C14` | `=G13` |
| 3 cells in `D14:D20` | Total revenues | `D14` | `=L13` |
| `B7` | Subject | `B7` | `=B2&" in "&B3` |
| `G11` | PY level | `G11` | `=0+B11*C11` |
| `L11` | AC level | `L11` | `=0+B11*D11` |
| `Y11` | span unit | `Y11` | `=MAX(AGGREGATE(4,6,I11:I30,J11:J30,K11:K30,N11:N30,O11:O30,P11:P30,E11:E30,Q11:Q30,R11:R30,S11:S30,T11:T30),-AGGREGATE(5,6,I11:I30,J11:J30,K11:K30,N11:N30,O11:O30,P11:P30,E11:E30,Q11:Q30,R11:R30,S11:S30,T11:T30),0.000000001)` |

**Charts**, in points from the top left of the sheet.

| Name | Type | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|---|
| `tier_wf_ac` | Stacked Bar | 1177.5 | 86 | 400 | 402 | 3 | -0.03 to 1.0556 |

| Chart | Series | Fill | Marker | Stem | Labelled |
|---|---|---|---|---|---|
| `tier_wf_ac` | `AC base` | no fill | None | - |  |
| `tier_wf_ac` | `AC adds` | solid #404040 | None | - | yes |
| `tier_wf_ac` | `AC subs` | solid #7F7F7F | None | - | yes |

---

*Generated by `ibcs_doc.py` from the workbook itself. Every formula, colour and measurement above was read back out of it.*
