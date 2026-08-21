# Building `IBCS Charts-Progressive - Complex Versions.xlsx` by hand in Excel

This is the same workbook, done manually - no macros, no add-ins, no code. Work through Part 1 once and you have the sheet every template in the library sits on; Part 2 builds the charts on top of it.

**Every range, formula, colour, gap width and axis bound below was read back out of `IBCS Charts-Progressive - Complex Versions.xlsx` after it was written.** Where this says a formula sits in `C03A!E13`, it does; where it says a fill is `#404040`, that is the value in the file. The prose was written once. The numbers are not typed twice, which is the only way a document like this stays true to the thing it describes.

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
| [C01A](#c01a) | a structure chart - net premiums written | 2 |
| [C02A](#c02a) | a structure chart - net premiums earned | 1 |
| [C03A](#c03a) | a tier stack - net premiums written | 3 |
| [C04A](#c04a) | a tier stack - underwriting result | 3 |
| [C05X](#c05x) | a tier stack - underwriting result | 3 |
| [C06F](#c06f) | a tier stack - net premiums written | 3 |
| [C07C](#c07c) | a line sheet - underwriting result | 1 |
| [C08H](#c08h) | a line sheet - loss and loss adjustment expense reserves, net | 3 |
| [C09C](#c09c) | an XY plot - net premiums earned | 1 |
| [C10D](#c10d) | an XY plot - underwriting performance by segment | 1 |
| [C11A](#c11a) | a driver tree - return on equity | 6 |
| [C12A](#c12a) | a tier stack - profit and loss statement | 4 |
| [C13D](#c13d) | a panel grid - net premiums written | 2 |
| [T01B](#t01b) | a table with drawn columns - underwriting result | 0 |
| [T02A](#t02a) | a table with drawn columns - underwriting result | 4 |
| [T03A](#t03a) | a table with drawn columns - underwriting result | 0 |
| [T04A](#t04a) | a table with drawn columns - underwriting result | 2 |

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

```
="Δ"&$B$5&"%"
```

Read from `C03A!B9`.

This is the part people skip, and it is the part that makes the sheet worth keeping. **Nothing a chart says may be typed into the chart.** Every text box on the finished chart is *linked* to one of these cells, so retyping `kEUR` as `kGBP` in `B3` changes the subject line and every caption at once. Type it into the chart and it is a lie the first time the data changes.

To link a text box: draw it, then with the box selected click in the **formula bar**, type `=` and click the cell. Not into the box - into the formula bar. It is the one place in Excel where that is the whole technique.

## Step 2 - Two zones, and nothing straddles them

The sheet is a **data zone** on the left of declared width, and the **charts** to its right. The print area is the chart zone alone - the data is the input, not the deliverable.

| Sheet | Cells used | Print area |
|---|---|---|
| Sources | `A1:D44` | `-` |
| C01A | `A1:F80` | `G1:P86` |
| C02A | `A1:V47` | `W1:AI13` |
| C03A | `A1:AC26` | `AD1:AO36` |
| C04A | `A1:Y31` | `Z1:AR30` |
| C05X | `A1:AU31` | `AV1:BI44` |
| C06F | `A1:AM31` | `AN1:BG34` |
| C07C | `A1:M59` | `N1:AA6` |
| C08H | `A1:U65` | `V1:AL6` |
| C09C | `A1:E146` | `F1:T42` |
| C10D | `A1:F19` | `F1:T37` |
| C11A | `A1:H62` | `I4:Z164` |
| C12A | `A1:AU32` | `AV1:BS33` |
| C13D | `A1:BU106` | `A109:N148` |
| T01B | `A1:O23` | `A9:O23` |
| T02A | `A1:AD23` | `A9:L23` |
| T03A | `A1:H24` | `A9:H24` |
| T04A | `A1:Q24` | `A9:G24` |

A **table** sheet is the other way round: on T01B, T02A, T03A, T04A the grid *is* the deliverable, so the print area is the grid, starting below the six typed title rows. Same rule - the print area holds what you are publishing and the typed inputs are not it.

> **Never AutoFit a column a chart is positioned against.** The charts are placed in points, measured from where the data zone ends. One long customer name and AutoFit moves that boundary, and the charts end up drawn on top of the source data. Set column widths by hand: **Home -> Format -> Column Width**.

## Step 3 - The data block

On `C03A` the headers are on row **12** (`A12:AC12`) and the data runs rows **13 to 24**.

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
=IF($E13>0,$E13,NA())
```

Read from `C03A!I13`.

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
=MAX(AGGREGATE(4,6,C13:C24,D13:D24,E13:E24,I13:I24,J13:J24),-AGGREGATE(5,6,C13:C24,D13:D24,E13:E24,I13:I24,J13:J24),0.000000001)
```

Read from `C03A!O13`. `AGGREGATE(4,6,...)` is MAX ignoring errors, which matters because half of these columns are deliberately `NA()`.

**2. A scaled copy of every column a chart reads**, which is what you actually plot:

```
=C13/$O$13
```

Read from `C03A!Q13`.

**3. Axis bounds divided by the same span.** If your measure tier was designed to run 0 to 230, you type `0` and `230/span` - or rather, you work out that number once and type the result.

Dividing both the values and the bounds by one number is a visual no-op at today's figures and follows the data at any others. And because zero divided by anything is zero, the zero line, the reference rules and every caption positioned from plot geometry stay exactly where they were.

Hide the scaled columns - **right-click the column headers -> Hide**. A hidden column measures zero width, so you can add all of this to a sheet without moving a single chart on it.

> **Where this cannot be used.** Every chart in this library hides its value axis, which is why dividing by a span is invisible. The two XY sheets are the exception - they show gridlines and tick labels, and dividing those by a span would print `0.25` where the data says `29.16%`. Those two get their axis fitted to the data once, at build time, and do not follow an edit. If you are building a chart whose axis a reader actually reads, do the same.

## Step 6 - A text column for every label

A data label linked to a cell shows the **cell's** value and ignores the label's own number format. Link a label to a raw variance and you get `31.61057692` printed beside a chart that says `+32`.

So the formatting happens in a cell:

```
=IF(ISNA(D13),"",TEXT(D13,"0"))
```

Read from `C03A!S13`.

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

`C03A` is net premiums written over 12 categories, in 3 tiers.

### 2.1 The columns, and what each is for

Headers on row **12**, data rows **13-24**. Typed cells are `C13:D24` - everything else is a formula.

| Column | Header | Formula in row 13 |
|---|---|---|
| A | Period | **typed** - `Jan`, and down |
| B | Scenario | **typed** - `AC`, and down |
| C | PY | **typed** - `6481`, and down |
| D | Measure | **typed** - `6735`, and down |
| E | ΔPY | `=D13-C13` |
| F | ΔPY% | `=IF(C13<=0,NA(),(D13-C13)/C13*100)` |
| G | Measure AC | `=IF($B13="AC",$D13,NA())` |
| H | Measure FC | `=IF($B13="AC",NA(),$D13)` |
| I | ΔPY up | `=IF($E13>0,$E13,NA())` |
| J | ΔPY down | `=IF($E13<0,$E13,NA())` |
| K | ΔPY% up | `=IF($F13>0,$F13,NA())` |
| L | ΔPY% down | `=IF($F13<0,$F13,NA())` |
| M | up length | `=IF($F13>0,$F13,0)` |
| N | down length | `=IF($F13<0,-$F13,0)` |
| O | span unit | `=MAX(AGGREGATE(4,6,C13:C24,D13:D24,E13:E24,I13:I24,J13:J24),-AGGREGATE(5,6,C13:C24,D13:D24,E13:E24,I13:I24,J13:J24),0.000000001)` |
| P | span rel | `=MAX(AGGREGATE(4,6,K13:K24,L13:L24,M13:M24,N13:N24),-AGGREGATE(5,6,K13:K24,L13:L24,M13:M24,N13:N24),0.000000001)` |
| Q | ref scaled | `=C13/$O$13` |
| R | measure scaled | `=D13/$O$13` |
| S | measure text | `=IF(ISNA(D13),"",TEXT(D13,"0"))` |
| T | var_abs scaled | `=E13/$O$13` |
| U | var_abs text | `=IF(ISNA(E13),"",TEXT(E13,"+0;-0"))` |
| V | abs_up scaled | `=I13/$O$13` |
| W | abs_dn scaled | `=J13/$O$13` |
| X | rel_up scaled | `=K13/$P$13` |
| Y | rel_up text | `=IF(ISNA(K13),"",TEXT(K13,"+0.0;-0.0"))` |
| Z | rel_dn scaled | `=L13/$P$13` |
| AA | rel_dn text | `=IF(ISNA(L13),"",TEXT(L13,"+0.0;-0.0"))` |
| AB | rel_up_len scaled | `=M13/$P$13` |
| AC | rel_dn_len scaled | `=N13/$P$13` |

Enter each formula in row 13 and fill down to row 24.

### 2.2 Check the numbers before you go near a chart

If these are right the chart cannot go far wrong; if they are wrong no amount of formatting will save it.

| Row | Period | Scenario | PY | Measure | ΔPY | ΔPY% | Measure AC | Measure FC |
|---|---|---|---|---|---|---|---|---|
| 13 | Jan | AC | 6481 | 6735 | `=D13-C13` | `=IF(C13<=0,NA(),(D13-C13)/C13*100)` | `=IF($B13="AC",$D13,NA())` | `=IF($B13="AC",NA(),$D13)` |
| 14 | Feb | AC | 6684 | 6995 | `=D14-C14` | `=IF(C14<=0,NA(),(D14-C14)/C14*100)` | `=IF($B14="AC",$D14,NA())` | `=IF($B14="AC",NA(),$D14)` |
| 15 | Mar | AC | 9041 | 9911 | `=D15-C15` | `=IF(C15<=0,NA(),(D15-C15)/C15*100)` | `=IF($B15="AC",$D15,NA())` | `=IF($B15="AC",NA(),$D15)` |
| 16 | Apr | AC | 6837 | 7278 | `=D16-C16` | `=IF(C16<=0,NA(),(D16-C16)/C16*100)` | `=IF($B16="AC",$D16,NA())` | `=IF($B16="AC",NA(),$D16)` |
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

### 2.5 The variance tier above it

Same construction, different scale, and it is a **separate chart object**. Insert it exactly as in 2.3, from the two variance columns you split in Part 1 step 4.

| | |
|---|---|
| Name | `tier_var_abs` |
| Type | Clustered Column |
| Size | 520 x 104 pt |
| Axis | -0.173 to 0.263 |
| Gap width | 47% |
| Overlap | 100% |

| Series | Fill |
|---|---|
| `ΔPY` | no fill |
| `ΔPY up` | solid #8CB400 |
| `ΔPY dn` | solid #FF0000 |

The invisible one is not a mistake. A variance drawn from a baseline that is not zero needs a transparent segment underneath holding the offset - the same trick a waterfall uses. Where the variance runs from zero, that series is still there and still zero, so the two tiers stay interchangeable.

### 2.6 The relative-variance tier: pins

A relative variance is drawn as a **pin** - a thin stem with a head marker - and there is no pin chart type. The obvious construction, a very narrow column, does not work: `Gap Width` caps at 500%, which on a twelve-category axis bottoms out around a 9px stem where the reference draws 5.

What works is a **line chart with markers, no line, and a custom Y error bar as the stem**. An error bar's weight is in *points*, so the stem is exactly as thin as you ask for.

1. **Insert -> Line with Markers** from the two split relative-variance columns.
2. Each series -> `Ctrl+1` -> Fill & Line -> **Line: No line**. The markers are the pins; the line between them means nothing.
3. Marker Options -> **Square**, size to taste, and set **Marker Fill** and **Marker Border** - not the series fill, which on a marker-only series is not what you are looking at. The head carries the *minuend's* scenario, which is what says whether a relative variance was measured or forecast.
4. With the series selected, **Chart Elements (+) -> Error Bars -> More Options**. Direction **Minus** for the series that goes up and **Plus** for the one that goes down, End Style **No Cap**, then **Custom -> Specify Value** and point *both* boxes at that series' stem-length column.
5. The error bar -> `Ctrl+1` -> Line -> width **3.8pt**.

| Series | Marker | Stem | Stem colour | Heads |
|---|---|---|---|---|
| `ΔPY% up` | Square | 3.8pt | #8CB400 | points 1-7: head #404040, border #404040; points 8-12: head #F2F2F2, border #404040 |
| `ΔPY% down` | Square | 3.8pt | #FF0000 | points 1-7: head #404040, border #404040; points 8-12: head #F2F2F2, border #404040 |

The two stems are different colours because the direction of a variance is its meaning, and an error bar takes one colour for a whole series - which is *why* there are two series here rather than one.

Four things about error bars are not in the documentation and each one fails quietly:

- The stem-length column must be **entirely numeric** - write `0` where the series has no point, never `NA()`. An error-bar range with one non-numeric cell is discarded whole.
- **End Style must be No Cap.** The cap is the little crossbar; with it on, your pin has a T on the end.
- The **direction** is the counter-intuitive one. An up-pin's stem runs *down* from the marker to zero, so it is a **Minus** error bar.
- On a line series, clicking a point and setting its **Line** colour styles the *connecting segment*, not the marker border. Use **Marker Border** for the border.

### 2.7 Line the tiers up - the step that makes it one figure

Three charts that nearly line up look like a mistake. They have to be exact, and you get there by typing numbers, not by dragging.

For each chart: `Ctrl+1` -> **Size & Properties**, and set Height, Width, and under Position the Horizontal and Vertical offsets:

| Chart | Left | Top | Width | Height |
|---|---|---|---|---|
| `tier_var_rel` | 683 | 96 | 520 | 104 |
| `tier_var_abs` | 683 | 204 | 520 | 104 |
| `tier_measure` | 683 | 312 | 520 | 244 |

That is not enough on its own. Two charts of the same width can still have plot areas of different widths, because Excel sizes the plot area around whatever labels it has to fit. So set the **plot area** too: click inside the plot (not the chart), `Ctrl+1`, and set its size and position:

| Chart | Plot left | Plot top | Plot width | Plot height |
|---|---|---|---|---|
| `tier_var_rel` | 52.1 | 6 | 452 | 77 |
| `tier_var_abs` | 52.1 | 6 | 452 | 84 |
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
| `tier_var_rel` / `ΔPY% up` | yes |
| `tier_var_rel` / `ΔPY% down` | yes |
| `tier_var_abs` / `ΔPY` | yes |
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

## Building a table with drawn columns, on `T04A`

Part 1 applies unchanged. What differs:

- **The grid is the deliverable.** There is no chart zone; the table itself is what prints, from row 14 down. Set the column widths by hand and the row heights to a fixed pitch, because the drawn columns are positioned against those rows.
- **Scenario notation lives in the borders.** A column header over an actual gets a bottom border in the scenario's own weight and colour; the double rule that says *plan* is Excel's own **xlDouble** border style - the outlined fill seen edge-on. Nothing here needs a drawn shape.
- **A drawn column is an ordinary variance chart** sized to the row pitch and positioned over the block it belongs to, with no axis, no gridlines and a transparent chart area. It is the tier stack of Part 2, one column wide.
- **The threshold is a cell.** Conditional formatting tests against the same cell the footnote displays, so the note and the colouring cannot disagree. IBCS requires the threshold to be stated wherever red is applied.
- **Subtotals are formulas**, never typed:

```
=SUM(B10:B14)
```

Read from `C01A!B15`.


**The charts on this sheet, as built:**

| Chart | Type | Left | Top | Width | Height | Series |
|---|---|---|---|---|---|---|
| `panel_dpl` | Clustered Bar | 375.5 | 538.9 | 106.2 | 123.8 | 3 |
| `panel_dplp` | Clustered Bar | 478 | 538.9 | 106.2 | 123.8 | 3 |

---

## Building a structure chart, on `C01A`

Part 1 applies unchanged. What differs:

- **One data block per panel**, side by side, and every panel on **one shared scale** - which is the entire point of the template. Work out the tallest column across all panels, round it up, and give every panel that maximum. Do not let Excel scale each one.
- **A stacked category is not a stacked scenario.** Where the bands are business areas, every band is an actual, so the fill cannot carry the scenario any more. The *column* carries it - a plan column is outlined - and the bands take a light-to-dark ramp.
- **Insert as a Stacked Column (or Bar)**, one series per band, bottom band first.
- **Small bands get no label.** Below about 3%% of the axis a number does not fit, and a label that overlaps its neighbour is worse than an absent one. That test belongs in the label formula, not in your judgement:

```
=IF(ISNA(D13),"",TEXT(D13,"0"))
```

Read from `C03A!S13`.


**The charts on this sheet, as built:**

| Chart | Type | Left | Top | Width | Height | Series |
|---|---|---|---|---|---|---|
| `panel_segment` | chart type -4111 | 407.5 | 81 | 300 | 430 | 6 |
| `panel_state` | chart type -4111 | 731.5 | 81 | 96 | 430 | 7 |

---

## Building a line sheet, on `C07C`

Part 1 applies unchanged. What differs:

- **The data runs in rows, not columns** - one row per series across the periods. Everything in Part 1 still applies, transposed.
- **A line must show its markers.** The line between two points is a connector, not data - there is no value at half past March - so the marker is what says measured, planned or expected. That is how one line changes from actual to forecast part way along, and it is why you cannot use a plain line chart here.
- **It is a combo chart**: columns for the monthly tier, lines for the cumulative. **Chart Design -> Change Chart Type -> Combo**, and set the type per series.
- **A flow and a level cannot share an x position.** A stock is drawn at period *boundaries* - an opening balance and one closing per period - while the movements that produce it sit inside the periods they belong to. Give them the same x and neither means anything.

**The charts on this sheet, as built:**

| Chart | Type | Left | Top | Width | Height | Series |
|---|---|---|---|---|---|---|
| `line_chart` | chart type -4111 | 644 | 81 | 620 | 444.5 | 7 |

---

## Building an XY plot, on `C09C`

Part 1 applies unchanged. What differs:

- **This is the one family that shows its value axis**, so the scale block of Part 1 step 5 **cannot be used**. Divide by a span here and the axis prints `0.25` where your data says `29.16%`. Fit the axis to the data once instead, and accept that it will not follow an edit.
- **Insert -> Scatter, or Bubble.** Each category that needs its own colour is its own series, because a series takes one colour.
- **A mark's area carries the value; its radius never does.** Bubble area goes as the value, so the radius goes as its square root. Excel's *Represent bubble size as: Area* does this for you - check it is set, because *Width* is also on that menu and it overstates a large value by the square of the ratio.
- **A colour that carries a category needs a key; a scenario never does.** Solid dark means measured on every IBCS page ever printed. A product line means nothing outside its own chart, so it takes an accent colour - deliberately not one of the scenario greys - and a legend. The presence of the key is itself the signal that the colour is not notation.
- **Where marks overlap, paint order is notation.** The mark the reader is being asked about goes last, even where it is the smaller one. Reorder with **Select Data -> the up/down arrows**.

**The charts on this sheet, as built:**

| Chart | Type | Left | Top | Width | Height | Series |
|---|---|---|---|---|---|---|
| `xy_C09C` | Scatter (X Y) | 310.5 | 99 | 700 | 560 | 6 |

---

## Building a driver tree, on `C11A`

Part 1 applies unchanged. What differs:

- **Each box is its own small chart**, and the arithmetic between them is text boxes and lines. There is no tree chart type and there is no way to fake one.
- **Boxes that share a unit share a scale.** Return, net sales and capital are all kEUR, so all three get the same points-per-unit - otherwise the tree invites a comparison it does not support.
- **The connectors state a calculation.** `ROS x Turnover = ROI` is written on the page, so the reader can check the tree against itself.
- Group the whole thing (**select all -> right-click -> Group**) once it is right, or moving it will take it apart.

**The charts on this sheet, as built:**

| Chart | Type | Left | Top | Width | Height | Series |
|---|---|---|---|---|---|---|
| `tree_roi` | Clustered Column | 441.5 | 696.2 | 232 | 534.1 | 1 |
| `tree_ros` | Clustered Column | 715.5 | 115 | 232 | 273.4 | 1 |
| `tree_turnover` | Clustered Column | 715.5 | 394.4 | 232 | 1420.9 | 1 |
| `tree_return` | Clustered Column | 989.5 | 451.6 | 232 | 132 | 1 |
| `tree_net_sales` | Clustered Column | 989.5 | 588.3 | 232 | 617.6 | 1 |
| `tree_capital` | Clustered Column | 989.5 | 1209 | 232 | 273.2 | 1 |

---

## Building a panel grid, on `C13D`

Part 1 applies unchanged. What differs:

- **Do not build this by copying one chart R x C times.** That is the usual advice and it leaves you maintaining a dozen charts whose axes drift apart the first time the data changes.
- The whole grid is **one chart object**, with the value axis divided into *bands* - one per grid row - and each panel's data arithmetically squeezed into its band:

```
plotted = band + (value - vmin) / (vmax - vmin) * band_frac
```

- Everything that is not data - the dividers, the panel titles, the tick labels - is drawn by extra XY scatter series pretending to be chart furniture, on the **secondary** axes.
- This one is genuinely hard to do by hand. The companion `panel-charts` skill builds it, and its `references/manual-steps.md` is a walkthrough of its own; that is the right document for this family rather than a paragraph here.

**The charts on this sheet, as built:**

| Chart | Type | Left | Top | Width | Height | Series |
|---|---|---|---|---|---|---|
| `panel_grid` | chart type -4111 | 0 | 1646.5 | 720 | 500 | 13 |
| `reference_panel` | Clustered Column | 549.9 | 2025.3 | 155.1 | 95.9 | 1 |

---

# Part 3 - Sheet by sheet

What is on each sheet, for rebuilding one template rather than learning the method.

### Sources

> AC

**A tier stack.**

| | |
|---|---|
| Cells used | `A1:D44` |
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
| Cells used | `A1:F80` |
| Print area | `G1:P86` |
| Formula cells | 82, in 7 shapes |
| Typed cells | 31 |

**Typed values** - 31 cells in `B10:F23`. Everything else is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in that range, so go by the shading.

**Every formula on the sheet.** Cells that say the same thing about their own position are one row here, with the range they cover and the first of them written out.

| Cells | Named | Example | Formula |
|---|---|---|---|
| 37 cells in `B38:F52` | - | `B38` | `=IF(ISBLANK(B10),NA(),B10/$B$26)` |
| 31 cells in `B66:F80` | - | `B66` | `=IF(ISBLANK(B10),"",IF(B10/83180*264<(LEN(TEXT(B10,"# ##0")))*5,"",TEXT(B10,"# ##0")))` |
| 6 cells in `B70:F79` | - | `B70` | `=IF(ISBLANK(B14),"",IF(B14/83180*264<(LEN(TEXT(B14,"# ##0"))+LEN(TEXT(B15,"# ##0")))*5,"",TEXT(B14,"# ##0")))` |
| `B15:F15` | Total | `B15` | `=SUM(B10:B14)` |
| `B7` | Subject | `B7` | `=B2&IF(B3="",""," in "&B3)` |
| `B24` | Total | `B24` | `=SUM(B18:B23)` |
| `B26` | span | `B26` | `=MAX(AGGREGATE(4,6,B15:F15,B24:B24),0.000000001)` |

**Charts**, in points from the top left of the sheet.

| Name | Type | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|---|
| `panel_segment` | chart type -4111 | 407.5 | 81 | 300 | 430 | 6 | 0 to 1.0001 |
| `panel_state` | chart type -4111 | 731.5 | 81 | 96 | 430 | 7 | 0 to 1.0001 |

| Chart | Series | Fill | Marker | Stem | Labelled |
|---|---|---|---|---|---|
| `panel_segment` | `Agency auto` | solid #404040 | None | - | yes |
| `panel_segment` | `Direct auto` | solid #595959 | None | - | yes |
| `panel_segment` | `Property` | solid #7F7F7F | None | - | yes |
| `panel_segment` | `Commercial` | solid #BFBFBF | None | - | yes |
| `panel_segment` | `Other` | solid #D9D9D9 | None | - | yes |
| `panel_segment` | `Total` | solid #000000 | None | - | yes |
| `panel_state` | `Texas` | solid #404040 | None | - | yes |
| `panel_state` | `Florida` | solid #595959 | None | - | yes |
| `panel_state` | `California` | solid #7F7F7F | None | - | yes |
| `panel_state` | `Georgia` | solid #A6A6A6 | None | - | yes |
| `panel_state` | `New York` | solid #BFBFBF | None | - | yes |
| `panel_state` | `All other` | solid #D9D9D9 | None | - | yes |
| `panel_state` | `Total` | solid #000000 | None | - | yes |

### C02A

> Of 81,659 $m earned in 2025, 53,866 went to losses and 17,523 to expenses, leaving an underwriting result of 10,270 $m (12.6%); Property keeps the largest share of its premium at 24.9% and Direct auto the smallest at 9.9%

**A structure chart.**

| | |
|---|---|
| Cells used | `A1:V47` |
| Print area | `W1:AI13` |
| Formula cells | 62, in 8 shapes |
| Typed cells | 12 |

**Typed values** - 12 cells in `B10:F12`. Everything else is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in that range, so go by the shading.

**Every formula on the sheet.** Cells that say the same thing about their own position are one row here, with the range they cover and the first of them written out.

| Cells | Named | Example | Formula |
|---|---|---|---|
| `B27:G30` | - | `B27` | `=IF(ISBLANK(B10),NA(),B10/$B$15)` |
| 18 cells in `B44:G47` | - | `B44` | `=IF(ISBLANK(B10),"",IF(B10/38300*492.8<(LEN(TEXT(B10,"# ##0")))*5,"",TEXT(B10,"# ##0")))` |
| `B13:G13` | Total | `B13` | `=SUM(B10:B12)` |
| `B46:G46` | - | `B46` | `=IF(ISBLANK(B12),"",IF(B12/38300*492.8<(LEN(TEXT(B12,"# ##0"))+LEN(TEXT(B13,"# ##0")))*5,"",TEXT(B12,"# ##0")))` |
| `E10:E12` | Personal Lines | `E10` | `=SUM(B10:D10)` |
| `G10:G12` | Companywide | `G10` | `=SUM(B10:D10)+F10` |
| `B7` | Subject | `B7` | `=B2&IF(B3="",""," in "&B3)` |
| `B15` | span | `B15` | `=MAX(AGGREGATE(4,6,$B$13:$D$13,$F$13:$F$13),0.000000001)` |

**Charts**, in points from the top left of the sheet.

| Name | Type | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|---|
| `panel_premium` | Stacked Bar | 1175.5 | 81 | 560 | 148.6 | 4 | 0 to 1.0005 |

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
| Cells used | `A1:AC26` |
| Print area | `AD1:AO36` |
| Hidden scale columns | `O` |
| Formula cells | 281, in 25 shapes |
| Typed cells | 24 |

**Typed values** - `C13:D24`. Everything else is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in that range, so go by the shading.

**Every formula on the sheet.** Cells that say the same thing about their own position are one row here, with the range they cover and the first of them written out.

| Cells | Named | Example | Formula |
|---|---|---|---|
| `Q13:R24` | ref scaled, measure scaled | `Q13` | `=C13/$O$13` |
| `V13:W24` | abs_up scaled, abs_dn scaled | `V13` | `=I13/$O$13` |
| `AB13:AC24` | rel_up_len scaled, rel_dn_len scaled | `AB13` | `=M13/$P$13` |
| `E13:E24` | ΔPY | `E13` | `=D13-C13` |
| `F13:F24` | ΔPY% | `F13` | `=IF(C13<=0,NA(),(D13-C13)/C13*100)` |
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

**Charts**, in points from the top left of the sheet.

| Name | Type | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|---|
| `tier_var_rel` | Line with Markers | 683 | 96 | 520 | 104 | 2 | -0.8671 to 1.3295 |
| `tier_var_abs` | Clustered Column | 683 | 204 | 520 | 104 | 3 | -0.173 to 0.263 |
| `tier_measure` | Clustered Column | 683 | 312 | 520 | 244 | 2 | 0 to 1.09 |

| Chart | Series | Fill | Marker | Stem | Labelled |
|---|---|---|---|---|---|
| `tier_var_rel` | `ΔPY% up` | points 1-7: head #404040, border #404040; points 8-12: head #F2F2F2, border #404040 | Square | 3.8pt #8CB400 | yes |
| `tier_var_rel` | `ΔPY% down` | points 1-7: head #404040, border #404040; points 8-12: head #F2F2F2, border #404040 | Square | 3.8pt #FF0000 | yes |
| `tier_var_abs` | `ΔPY` | no fill | None | - | yes |
| `tier_var_abs` | `ΔPY up` | points 1-7: solid #8CB400; points 8-12: Wide upward diagonal, #8CB400 on #FFFFFF | None | - |  |
| `tier_var_abs` | `ΔPY dn` | points 1-7: solid #FF0000; points 8-12: Wide upward diagonal, #FF0000 on #FFFFFF | None | - |  |
| `tier_measure` | `PY` | solid #A6A6A6 | None | - |  |
| `tier_measure` | `AC` | points 1-7: solid #404040; points 8-12: Wide upward diagonal, #404040 on #F2F2F2 | None | - | yes |

### C04A

> 6 of 12 months beat the 96 combined ratio target in 2023; Dec by +670 $m and Mar missed it by 441 $m, a spread of 1,111 $m across the year

**A tier stack.**

| | |
|---|---|
| Cells used | `A1:Y31` |
| Print area | `Z1:AR30` |
| Hidden scale columns | `M` |
| Formula cells | 233, in 21 shapes |
| Typed cells | 24 |

**Typed values** - `C13:D24`. Everything else is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in that range, so go by the shading.

**Every formula on the sheet.** Cells that say the same thing about their own position are one row here, with the range they cover and the first of them written out.

| Cells | Named | Example | Formula |
|---|---|---|---|
| 36 cells in `P13:U24` | measure scaled, abs_up scaled, abs_dn scaled | `P13` | `=C13/$M$13` |
| `X13:Y24` | rel_up scaled, rel_dn scaled | `X13` | `=I13/$N$13` |
| `E13:E24` | PL | `E13` | `=C13-D13` |
| `F13:F24` | ΔPL% | `F13` | `=IF(E13<=0,NA(),D13/E13*100)` |
| `G13:G24` | ΔPL up | `G13` | `=IF($D13>0,$D13,NA())` |
| `H13:H24` | ΔPL down | `H13` | `=IF($D13<0,$D13,NA())` |
| `I13:I24` | ΔPL% up | `I13` | `=IF($F13>0,$F13,NA())` |
| `J13:J24` | ΔPL% down | `J13` | `=IF($F13<0,$F13,NA())` |
| `K13:K24` | up length | `K13` | `=IF($F13>0,$F13,0)` |
| `L13:L24` | down length | `L13` | `=IF($F13<0,-$F13,0)` |
| `O13:O24` | ref scaled | `O13` | `=E13/$M$13` |
| `Q13:Q24` | measure text | `Q13` | `=IF(ISNA(C13),"",TEXT(C13,"0"))` |
| `R13:R24` | var_abs scaled | `R13` | `=D13/$M$13` |
| `S13:S24` | var_abs text | `S13` | `=IF(ISNA(D13),"",TEXT(D13,"+0;-0"))` |
| `V13:V24` | var_rel scaled | `V13` | `=F13/$N$13` |
| `W13:W24` | var_rel text | `W13` | `=IF(ISNA(F13),"",TEXT(F13,"+0;-0"))` |
| `B7` | Subject | `B7` | `=B2&" in "&B3` |
| `B9` | var_abs caption | `B9` | `="Δ"&$B$5&" "&$B$3` |
| `B10` | var_rel caption | `B10` | `="Δ"&$B$5&"%"` |
| `M13` | span unit | `M13` | `=MAX(AGGREGATE(4,6,E13:E31,C13:C31,D13:D31,G13:G31,H13:H31),-AGGREGATE(5,6,E13:E31,C13:C31,D13:D31,G13:G31,H13:H31),0.000000001)` |
| `N13` | span rel | `N13` | `=MAX(AGGREGATE(4,6,F13:F31,I13:I31,J13:J31),-AGGREGATE(5,6,F13:F31,I13:I31,J13:J31),0.000000001)` |

**Charts**, in points from the top left of the sheet.

| Name | Type | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|---|
| `tier_measure` | Clustered Bar | 631.5 | 86 | 430 | 400 | 2 | -0.35 to 1.826 |
| `tier_var_abs` | Clustered Bar | 1066.5 | 86 | 250 | 400 | 3 | -0.55 to 0.81 |
| `tier_var_rel` | Clustered Bar | 1321.5 | 86 | 170 | 400 | 3 | -0.95 to 1.15 |

| Chart | Series | Fill | Marker | Stem | Labelled |
|---|---|---|---|---|---|
| `tier_measure` | `PL` | solid #FFFFFF | None | - |  |
| `tier_measure` | `AC` | solid #404040 | None | - | yes |
| `tier_var_abs` | `ΔPL` | no fill | None | - | yes |
| `tier_var_abs` | `ΔPL up` | solid #8CB400 | None | - |  |
| `tier_var_abs` | `ΔPL dn` | solid #FF0000 | None | - |  |
| `tier_var_rel` | `ΔPL%` | no fill | None | - | yes |
| `tier_var_rel` | `ΔPL% up` | solid #8CB400 | None | - |  |
| `tier_var_rel` | `ΔPL% dn` | solid #FF0000 | None | - |  |

### C05X

> The underwriting result is running +1,200 $m (+11.7%) ahead of 2025, with 7 of twelve months up and Sep contributing +985 on its own. Both years sit far above the 96 combined ratio target, drawn as the second opening column

**A tier stack.**

| | |
|---|---|
| Cells used | `A1:AU31` |
| Print area | `AV1:BI44` |
| Hidden scale columns | `W` |
| Formula cells | 689, in 61 shapes |
| Typed cells | 25 |

**Typed values** - 25 cells in `D13:E25`. Everything else is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in that range, so go by the shading.

**Every formula on the sheet.** Cells that say the same thing about their own position are one row here, with the range they cover and the first of them written out.

| Cells | Named | Example | Formula |
|---|---|---|---|
| 51 cells in `AC12:AH28` | bar_ac scaled, wf_base scaled, wf_up_ac scaled | `AC12` | `=J12/$W$12` |
| 40 cells in `G12:T28` | ΔPY%, PY bar, AC bar, FC bar | `G12` | `=NA()` |
| 34 cells in `AE12:AJ28` | bar_fc scaled, wf_dn_ac scaled | `AE12` | `=K12/$W$12` |
| `AT12:AU28` | rel_up_len scaled, rel_dn_len scaled | `AT12` | `=U12/$X$12` |
| 33 cells in `F12:V28` | PY, level, base, up AC | `F12` | `=0` |
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
| `L27:L28` | vs PY | `L27` | `=$D$26` |
| `O27:O28` | vs PY | `O27` | `=IF(E27>0,E27,0)` |

*21 further named shapes are not listed - this sheet has 61 in all.*

**Charts**, in points from the top left of the sheet.

| Name | Type | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|---|
| `tier_var_rel` | Line with Markers | 1038.5 | 96 | 620 | 130 | 2 | -1.1 to 0.1136 |
| `tier_wf` | Stacked Column | 1038.5 | 236 | 620 | 110 | 5 | 0.7943 to 1.101 |
| `tier_measure` | Clustered Column | 1038.5 | 356 | 620 | 330 | 4 | -0.0024 to 1.02 |

| Chart | Series | Fill | Marker | Stem | Labelled |
|---|---|---|---|---|---|
| `tier_var_rel` | `ΔPY% up` | points 1-7: head #404040, border #404040; points 8-12: head #F2F2F2, border #404040; points 13-17: head #156082, border #FEFFFF | Square | 3.8pt #8CB400 | yes |
| `tier_var_rel` | `ΔPY% down` | points 1-7: head #404040, border #404040; points 8-12: head #F2F2F2, border #404040; points 13-17: head #FFFFFF, border #FFFFFF | Square | 3.8pt #FF0000 | yes |
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
| Cells used | `A1:AM31` |
| Print area | `AN1:BG34` |
| Hidden scale columns | `T` |
| Formula cells | 453, in 42 shapes |
| Typed cells | 24 |

**Typed values** - 24 cells in `D13:E25`. Everything else is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in that range, so go by the shading.

**Every formula on the sheet.** Cells that say the same thing about their own position are one row here, with the range they cover and the first of them written out.

| Cells | Named | Example | Formula |
|---|---|---|---|
| 55 cells in `F12:Q25` | PY, AC bar, PY short, PL bar | `F12` | `=0` |
| 42 cells in `AA12:AF25` | bar_py scaled, wf_base scaled, wf_up scaled | `AA12` | `=K12/$T$12` |
| `X12:Y25` | bar_short scaled, bar_pl scaled | `X12` | `=I12/$T$12` |
| 28 cells in `AC12:AH25` | bar_ac scaled, wf_dn scaled | `AC12` | `=L12/$T$12` |
| `AL12:AM25` | rel_up scaled, rel_dn scaled | `AL12` | `=R12/$U$12` |
| `V12:V25` | bar_state scaled | `V12` | `=H12/$T$12` |
| `W12:W25` | bar_state text | `W12` | `=IF(ISNA(H12),"",TEXT(H12,"#,##0"))` |
| `Z12:Z25` | bar_pl text | `Z12` | `=IF(ISNA(J12),"",TEXT(J12,"#,##0"))` |
| `AB12:AB25` | bar_py text | `AB12` | `=IF(ISNA(K12),"",TEXT(K12,"#,##0"))` |
| `AD12:AD25` | bar_ac text | `AD12` | `=IF(ISNA(L12),"",TEXT(L12,"#,##0"))` |
| `AG12:AG25` | wf_up text | `AG12` | `=IF(ISNA(P12),"",TEXT(P12,"+#,##0"))` |
| `AI12:AI25` | wf_dn text | `AI12` | `=IF(ISNA(Q12),"",TEXT(Q12,"""-""#,##0"))` |
| `AJ12:AJ25` | var_rel scaled | `AJ12` | `=E12/$U$12` |
| `AK12:AK25` | var_rel text | `AK12` | `=IF(ISNA(E12),"",TEXT(E12,"+0;-0"))` |
| 12 cells in `P13:P25` | Texas | `P13` | `=IF(G13>0,G13,0)` |
| 12 cells in `Q13:Q25` | Texas | `Q13` | `=IF(G13<0,-G13,0)` |
| `R13:R24` | Texas | `R13` | `=IF(E13>0,E13,NA())` |
| `S13:S24` | Texas | `S13` | `=IF(E13<0,E13,NA())` |
| `F13:F23` | Texas | `F13` | `=IF(E13=-100,NA(),D13/(1+E13/100))` |
| `G13:G23` | Texas | `G13` | `=D13-F13` |
| `H13:H23` | Texas | `H13` | `=D13` |
| `I13:I23` | Texas | `I13` | `=MAX(F13-D13,0)` |
| `M13:M23` | Texas | `M13` | `=N13+G13` |
| `N13:N23` | Texas | `N13` | `=M12` |
| `O13:O23` | Texas | `O13` | `=MIN(N13,M13)` |
| 4 cells in `R12:S25` | ΔPY% up, ΔPY% down | `R12` | `=NA()` |
| 2 cells in `N12:N24` | from | `N12` | `=M12` |
| `B7` | Subject | `B7` | `=B2&" in "&B3` |
| `B9` | var_rel caption | `B9` | `="Δ"&$B$5&"%"` |
| `D12` | AC | `D12` | `=SUM(F13:F23)` |
| `K12` | PY bar | `K12` | `=D12` |
| `M12` | level | `M12` | `=D12` |
| `T12` | span unit | `T12` | `=MAX(AGGREGATE(4,6,H12:H31,I12:I31,J12:J31,K12:K31,L12:L31,O12:O31,P12:P31,Q12:Q31),-AGGREGATE(5,6,H12:H31,I12:I31,J12:J31,K12:K31,L12:L31,O12:O31,P12:P31,Q12:Q31),0.000000001)` |
| `U12` | span rel | `U12` | `=MAX(AGGREGATE(4,6,E12:E31,R12:R31,S12:S31),-AGGREGATE(5,6,E12:E31,R12:R31,S12:S31),0.000000001)` |
| `D24` | AC | `D24` | `=SUM(D13:D23)` |
| `G24` | AC | `G24` | `=D24-$D$12` |
| `L24` | AC | `L24` | `=D24` |
| `M24` | AC | `M24` | `=M23` |
| `G25` | ΔPY | `G25` | `=$D$24-$D$12` |
| `M25` | ΔPY | `M25` | `=$D$24` |

*2 further named shapes are not listed - this sheet has 42 in all.*

**Charts**, in points from the top left of the sheet.

| Name | Type | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|---|
| `tier_measure` | Stacked Bar | 970.5 | 86 | 560 | 430 | 5 | 0 to 1.02 |
| `tier_wf` | Stacked Bar | 1540.5 | 86 | 140 | 430 | 3 | 0.8083 to 1.0865 |
| `tier_var_rel` | Clustered Bar | 1690.5 | 86 | 200 | 430 | 3 | 0 to 1.1 |

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
| `tier_var_rel` | `ΔPY%` | no fill | None | - | yes |
| `tier_var_rel` | `ΔPY% up` | solid #8CB400 | None | - |  |
| `tier_var_rel` | `ΔPY% dn` | solid #FF0000 | None | - |  |

### C07C

> Through Jul the underwriting result is +4,557 $m against the 96 combined ratio target - a 13.1% margin where the target implies 4.0%. Holding that margin closes the year +7,967 $m ahead of target, and the moving annual total has stayed above 10,000 $m all year

**A line sheet.**

| | |
|---|---|
| Cells used | `A1:M59` |
| Print area | `N1:AA6` |
| Formula cells | 183, in 22 shapes |
| Typed cells | 36 |

**Typed values** - 36 cells in `B11:M17`. Everything else is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in that range, so go by the shading.

**Every formula on the sheet.** Cells that say the same thing about their own position are one row here, with the range they cover and the first of them written out.

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
| `B7` | Subject | `B7` | `=B2&IF(B3="",""," in "&B3)` |
| `I14` | PL cumulative | `I14` | `=SUM(B11:I11)` |
| `J14` | PL cumulative | `J14` | `=SUM(B11:J11)` |
| `K14` | PL cumulative | `K14` | `=SUM(B11:K11)` |
| `L14` | PL cumulative | `L14` | `=SUM(B11:L11)` |
| `M14` | PL cumulative | `M14` | `=SUM(B11:M11)` |
| `H16` | FC cumulative | `H16` | `=SUM(B12:H12)+SUM(B13:H13)` |
| `I16` | FC cumulative | `I16` | `=SUM(B12:I12)+SUM(B13:I13)` |
| `J16` | FC cumulative | `J16` | `=SUM(B12:J12)+SUM(B13:J13)` |
| `K16` | FC cumulative | `K16` | `=SUM(B12:K12)+SUM(B13:K13)` |
| `L16` | FC cumulative | `L16` | `=SUM(B12:L12)+SUM(B13:L13)` |
| `M16` | FC cumulative | `M16` | `=SUM(B12:M12)+SUM(B13:M13)` |
| `B19` | span | `B19` | `=MAX(AGGREGATE(4,6,B11:M11,B12:M12,B13:M13,B14:M14,B15:M15,B16:M16,B17:M17),-AGGREGATE(5,6,B11:M11,B12:M12,B13:M13,B14:M14,B15:M15,B16:M16,B17:M17),0.000000001)` |

**Charts**, in points from the top left of the sheet.

| Name | Type | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|---|
| `line_chart` | chart type -4111 | 644 | 81 | 620 | 444.5 | 7 | 0 to 1.0387 |

| Chart | Series | Fill | Marker | Stem | Labelled |
|---|---|---|---|---|---|
| `line_chart` | `PL month` | solid #FFFFFF | None | - |  |
| `line_chart` | `AC month` | solid #404040 | None | - | yes |
| `line_chart` | `FC month` | Wide upward diagonal, #404040 on #F2F2F2 | None | - | yes |
| `line_chart` | `PL cumulative` | head #FFFFFF, border #404040 | Square | - | yes |
| `line_chart` | `AC cumulative` | head #404040, border #000000 | Square | - | yes |
| `line_chart` | `FC cumulative` | head #F2F2F2, border #404040 | Square | - | yes |
| `line_chart` | `MAT` | head #404040, border #000000 | Square | - | yes |

### C08H

> Reserves built +12,301 $m over 10 quarters - losses incurred ran ahead of claims paid in 10 of them. Progressive is running a 66.2 loss ratio where its 96 combined ratio target implies 71.9, so holding the target builds reserves faster rather than slower: 56,535 $m by end 2028

**A line sheet.**

| | |
|---|---|
| Cells used | `A1:U65` |
| Print area | `V1:AL6` |
| Formula cells | 262, in 30 shapes |
| Typed cells | 41 |

**Typed values** - 41 cells in `A12:U15`. Everything else is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in that range, so go by the shading.

**Every formula on the sheet.** Cells that say the same thing about their own position are one row here, with the range they cover and the first of them written out.

| Cells | Named | Example | Formula |
|---|---|---|---|
| 100 cells in `B35:U42` | Increase scaled | `B35` | `=IF(ISNA(B12),NA(),IF(ISBLANK(B12),NA(),B12/$B$21))` |
| 40 cells in `B58:U65` | Increase text | `B58` | `=IF(ISNA(B12),"",IF(ISBLANK(B12),"",TEXT(B12,"0")))` |
| 38 cells in `B16:U18` | Inventory AC | `M16` | `=NA()` |
| `B14:U14` | Inventory change | `B14` | `=B12-B13` |
| `B19:U19` | Decrease (drawn) | `B19` | `=-B13` |
| `B16:L16` | Inventory AC | `B16` | `=B15` |
| `N18:U18` | Inventory PL | `N18` | `=N15` |
| `L17:N17` | Inventory FC | `L17` | `=L15` |
| `B7` | Subject | `B7` | `=B2&IF(B3="",""," in "&B3)` |
| `B21` | span | `B21` | `=MAX(AGGREGATE(4,6,B16:U16,B17:U17,B18:U18,B12:U12,B19:U19),-AGGREGATE(5,6,B16:U16,B17:U17,B18:U18,B12:U12,B19:U19),0.000000001)` |

20 unnamed one-off formulas: `B15`, `C15`, `D15`, `E15`, `F15`, `G15`, `H15`, `I15`, `J15`, `K15`, `L15`, `M15`, ....

**Charts**, in points from the top left of the sheet.

| Name | Type | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|---|
| `c08_change` | Clustered Column | 865 | 81 | 760 | 179 | 1 | 0 to 2500 |
| `c08_level` | chart type 1 | 865 | 270 | 760 | 300 | 3 | -0.3636 to 1.0909 |
| `c08_flows` | Clustered Column | 865 | 270 | 760 | 300 | 2 | -0.3636 to 1.0909 |

| Chart | Series | Fill | Marker | Stem | Labelled |
|---|---|---|---|---|---|
| `c08_change` | `Inventory change` | points 1-10: solid #7F7F7F; points 13-20: solid #D9D9D9; points 11-12: Wide upward diagonal, #404040 on #F2F2F2 | None | - | yes |
| `c08_level` | `level AC` | solid #D9D9D9 | None | - |  |
| `c08_level` | `level FC` | Wide upward diagonal, #404040 on #F2F2F2 | None | - |  |
| `c08_level` | `level PL` | solid #F2F2F2 | None | - |  |
| `c08_flows` | `Increase` | points 1-10: solid #7F7F7F; points 13-20: solid #D9D9D9; points 11-12: Wide upward diagonal, #404040 on #F2F2F2 | None | - | yes |
| `c08_flows` | `Decrease drawn` | points 1-10: solid #404040; points 13-20: solid #F4F4F4; points 11-12: Wide upward diagonal, #404040 on #F2F2F2 | None | - | yes |

### C09C

> 17 of Agency auto's 23 months earned an underwriting result of 300 $m or more; the curves are constant result, so a point's distance beyond one is what its premium and its margin achieved together

**An xy plot.**

| | |
|---|---|
| Cells used | `A1:E146` |
| Print area | `F1:T42` |
| Formula cells | 251, in 6 shapes |
| Typed cells | 202 |

**Typed values** - 202 cells in `B11:E146`. Everything else is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in that range, so go by the shading.

**Every formula on the sheet.** Cells that say the same thing about their own position are one row here, with the range they cover and the first of them written out.

| Cells | Named | Example | Formula |
|---|---|---|---|
| 69 cells in `E11:E81` | - | `E11` | `=C11*B11/100` |
| `C87:C146` | - | `C87` | `=IF(100*100/$B87>3500,NA(),100*100/$B87)` |
| `D87:D146` | - | `D87` | `=IF(200*100/$B87>3500,NA(),200*100/$B87)` |
| `E87:E146` | - | `E87` | `=IF(300*100/$B87>3500,NA(),300*100/$B87)` |
| `B7` | Subject | `B7` | `=B2&IF(B3="",""," in "&B3)` |
| `E83` | Agency auto at 300 mUSD or more | `E83` | `=SUMPRODUCT((D11:D81="Agency auto")*(E11:E81>=300))` |

**Charts**, in points from the top left of the sheet.

| Name | Type | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|---|
| `xy_C09C` | Scatter (X Y) | 310.5 | 99 | 700 | 560 | 6 | 0 to 3500 |

| Chart | Series | Fill | Marker | Stem | Labelled |
|---|---|---|---|---|---|
| `xy_C09C` | `Agency auto` | head #404040, border #404040 | Circle | - |  |
| `xy_C09C` | `Direct auto` | head #FF6600, border #404040 | Circle | - |  |
| `xy_C09C` | `Commercial` | head #FFCE59, border #404040 | Circle | - |  |
| `xy_C09C` | `100 mUSD` | solid #000000 | None | - |  |
| `xy_C09C` | `200 mUSD` | solid #000000 | None | - |  |
| `xy_C09C` | `300 mUSD` | solid #000000 | None | - |  |

### C10D

> Property carries the highest combined ratio at 96.2 despite the lowest loss ratio in the group at 63.2; its expense ratio of 33.0 is 14 points above the leanest segment

**An xy plot.**

| | |
|---|---|
| Cells used | `A1:F19` |
| Print area | `F1:T37` |
| Formula cells | 2, in 2 shapes |
| Typed cells | 21 |

**Typed values** - 21 cells in `B11:E19`. Everything else is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in that range, so go by the shading.

**Every formula on the sheet.** Cells that say the same thing about their own position are one row here, with the range they cover and the first of them written out.

| Cells | Named | Example | Formula |
|---|---|---|---|
| `B7` | Subject | `B7` | `=B2&IF(B3="",""," in "&B3)` |
| `E10` | PY | `E10` | `=B8&" in "&C8` |

**Charts**, in points from the top left of the sheet.

| Name | Type | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|---|
| `xy_C10D` | Bubble | 321.5 | 99 | 660 | 470 | 3 | 55 to 80 |

| Chart | Series | Fill | Marker | Stem | Labelled |
|---|---|---|---|---|---|
| `xy_C10D` | `PY` | solid #D9D9D9 | None | - | yes |
| `xy_C10D` | `AC` | solid #404040 | None | - | yes |
| `xy_C10D` | `ACQ` | solid #0064FF | None | - | yes |

### C11A

> Return on equity moved from 24.1% to 33.9% between 2019 and 2025, and the tree says which half did it: premium to equity barely moved (2.65 to 2.69) while the underwriting margin went from 9.1% to 12.6%

**A driver tree.**

| | |
|---|---|
| Cells used | `A1:H62` |
| Print area | `I4:Z164` |
| Formula cells | 123, in 15 shapes |
| Typed cells | 21 |

**Typed values** - `B13:H15`. Everything else is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in that range, so go by the shading.

**Every formula on the sheet.** Cells that say the same thing about their own position are one row here, with the range they cover and the first of them written out.

| Cells | Named | Example | Formula |
|---|---|---|---|
| `B57:H62` | return text | `B57` | `=IF(ISNA(B13),"",IF(ISBLANK(B13),"",TEXT(B13,"0.0")))` |
| `B35:H37` | return scaled | `B35` | `=IF(ISNA(B13),NA(),IF(ISBLANK(B13),NA(),B13/$B$22))` |
| 14 cells in `B38:H40` | ros scaled | `B38` | `=IF(ISNA(B16),NA(),IF(ISBLANK(B16),NA(),B16/$B$20))` |
| `B16:H16` | Underwriting margin | `B16` | `=B13/B14*100` |
| `B17:H17` | Premium to equity | `B17` | `=B14/B15` |
| `B18:H18` | Return on equity | `B18` | `=B13/B15*100` |
| `B39:H39` | turnover scaled | `B39` | `=IF(ISNA(B17),NA(),IF(ISBLANK(B17),NA(),B17/$B$21))` |
| `C11:H11` | AC | `C11` | `=IF(C10=B10,"",C10)` |
| `C12:G12` | Axis year | `C12` | `=""` |
| 2 cells in `B12:H12` | Axis year | `B12` | `=B9` |
| `B7` | Subject | `B7` | `=B2&IF(B3="",""," in "&B3)` |
| `B11` | AC | `B11` | `=B10` |
| `B20` | span percent | `B20` | `=MAX(AGGREGATE(4,6,B18:H18,B16:H16),-AGGREGATE(5,6,B18:H18,B16:H16),0.000000001)` |
| `B21` | span turnover | `B21` | `=MAX(AGGREGATE(4,6,B17:H17),-AGGREGATE(5,6,B17:H17),0.000000001)` |
| `B22` | span kEUR | `B22` | `=MAX(AGGREGATE(4,6,B13:H13,B14:H14,B15:H15),-AGGREGATE(5,6,B13:H13,B14:H14,B15:H15),0.000000001)` |

**Charts**, in points from the top left of the sheet.

| Name | Type | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|---|
| `tree_roi` | Clustered Column | 441.5 | 696.2 | 232 | 534.1 | 1 | -0.1474 to 1.1789 |
| `tree_ros` | Clustered Column | 715.5 | 115 | 232 | 273.4 | 1 | -0.1474 to 0.4421 |
| `tree_turnover` | Clustered Column | 715.5 | 394.4 | 232 | 1420.9 | 1 | -0.1614 to 1.2909 |
| `tree_return` | Clustered Column | 989.5 | 451.6 | 232 | 132 | 1 | -0.0177 to 0.1436 |
| `tree_net_sales` | Clustered Column | 989.5 | 588.3 | 232 | 617.6 | 1 | -0.14 to 1.14 |
| `tree_capital` | Clustered Column | 989.5 | 1209 | 232 | 273.2 | 1 | -0.052 to 0.4233 |

| Chart | Series | Fill | Marker | Stem | Labelled |
|---|---|---|---|---|---|
| `tree_roi` | `Return on equity` | solid #156082 | None | - | yes |
| `tree_ros` | `Underwriting margin` | solid #156082 | None | - | yes |
| `tree_turnover` | `Premium to equity` | solid #156082 | None | - | yes |
| `tree_return` | `Underwriting result` | solid #156082 | None | - | yes |
| `tree_net_sales` | `Net premiums earned` | solid #156082 | None | - | yes |
| `tree_capital` | `Shareholders' equity` | solid #156082 | None | - | yes |

### C12A

> Net income rose from 722 to 3,903 $m, +441%, as an underwriting result that had nearly vanished in 2022 recovered; 3 lines move too far to be drawn to scale and carry the overflow marks

**A tier stack.**

| | |
|---|---|
| Cells used | `A1:AU32` |
| Print area | `AV1:BS33` |
| Hidden scale columns | `AP`, `AU`, `Y` |
| Formula cells | 430, in 38 shapes |
| Typed cells | 24 |

**Typed values** - 24 cells in `B13:D22`. Everything else is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in that range, so go by the shading.

**Every formula on the sheet.** Cells that say the same thing about their own position are one row here, with the range they cover and the first of them written out.

| Cells | Named | Example | Formula |
|---|---|---|---|
| 40 cells in `Z13:AF22` | py_base scaled, py_add scaled, ac_base scaled, ac_add scaled | `Z13` | `=I13/$Y$13` |
| `AL13:AO22` | abs_pos_good scaled, abs_pos_bad scaled, abs_neg_good scaled, abs_neg_bad scaled | `AL13` | `=Q13/$Y$13` |
| `AQ13:AT22` | rel_pos_good clipped, rel_pos_bad clipped, rel_neg_good clipped, rel_neg_bad clipped | `AQ13` | `=IF(ISNA(U13),NA(),MEDIAN(-31.5,U13,238.5))` |
| 20 cells in `I13:N22` | PY base, AC base | `I13` | `=MIN(H13,G13)` |
| 20 cells in `AB13:AG22` | py_add text, ac_add text | `AB13` | `=IF(ISNA(J13),"",TEXT(J13,"#,##0"))` |
| 20 cells in `AC13:AH22` | py_sub scaled, ac_sub scaled | `AC13` | `=K13/$Y$13` |
| 20 cells in `AD13:AI22` | py_sub text, ac_sub text | `AD13` | `=IF(ISNA(K13),"",TEXT(K13,"#,##0"))` |
| 12 cells in `H14:M21` | Investment income | `H14` | `=G13` |
| `E13:E22` | ΔPY | `E13` | `=D13-C13` |
| `F13:F22` | ΔPY% | `F13` | `=IF(C13<=0,NA(),(D13-C13)/C13*100)` |
| `J13:J22` | PY adds | `J13` | `=IF(B13>0,ABS(G13-H13),0)` |
| `K13:K22` | PY subs | `K13` | `=IF(B13<0,ABS(G13-H13),0)` |
| `O13:O22` | AC adds | `O13` | `=IF(B13>0,ABS(L13-M13),0)` |
| `P13:P22` | AC subs | `P13` | `=IF(B13<0,ABS(L13-M13),0)` |
| `Q13:Q22` | ΔPY up ok | `Q13` | `=IF(AND($E13>0,$B13>0),$E13,NA())` |
| `R13:R22` | ΔPY up bad | `R13` | `=IF(AND($E13>0,$B13<0),$E13,NA())` |
| `S13:S22` | ΔPY dn ok | `S13` | `=IF(AND($E13<0,$B13<0),$E13,NA())` |
| `T13:T22` | ΔPY dn bad | `T13` | `=IF(AND($E13<0,$B13>0),$E13,NA())` |
| `U13:U22` | ΔPY% up ok | `U13` | `=IF(AND($F13>0,$B13>0),$F13,NA())` |
| `V13:V22` | ΔPY% up bad | `V13` | `=IF(AND($F13>0,$B13<0),$F13,NA())` |
| `W13:W22` | ΔPY% dn ok | `W13` | `=IF(AND($F13<0,$B13<0),$F13,NA())` |
| `X13:X22` | ΔPY% dn bad | `X13` | `=IF(AND($F13<0,$B13>0),$F13,NA())` |
| `AJ13:AJ22` | var_abs scaled | `AJ13` | `=E13/$Y$13` |
| `AK13:AK22` | var_abs text | `AK13` | `=IF(ISNA(E13),"",TEXT(E13,"+0;-0"))` |
| `AP13:AP22` | var_rel clipped | `AP13` | `=IF(ISNA(F13),NA(),MEDIAN(-31.5,F13,238.5))` |
| `AU13:AU22` | ΔPY% text | `AU13` | `=IF(ISNA(F13),"",IF(F13>265,TEXT(F13,"+0;-0")&REPT(UNICHAR(9658),MIN(3,MAX(1,INT(F13/265)))),IF(F13<-35,REPT(UNICHAR(9668),MIN(3,MAX(1,INT(-F13/35))))&TEXT(F13,"+0;-0"),TEXT(F13,"+0;-0"))))` |
| 8 cells in `H13:M22` | PY from, AC from | `H13` | `=0` |
| 6 cells in `G14:G21` | Investment income | `G14` | `=G13+B14*C14` |
| 6 cells in `L14:L21` | Investment income | `L14` | `=L13+B14*D14` |
| 6 cells in `G16:L22` | Total revenues | `G16` | `=G15` |
| 3 cells in `C16:C22` | Total revenues | `C16` | `=G15` |
| 3 cells in `D16:D22` | Total revenues | `D16` | `=L15` |
| `B7` | Subject | `B7` | `=B2&" in "&B3` |
| `B9` | var_abs caption | `B9` | `="Δ"&$B$5&" "&$B$3` |
| `B10` | var_rel caption | `B10` | `="Δ"&$B$5&"%"` |
| `G13` | PY level | `G13` | `=0+B13*C13` |
| `L13` | AC level | `L13` | `=0+B13*D13` |
| `Y13` | span unit | `Y13` | `=MAX(AGGREGATE(4,6,I13:I32,J13:J32,K13:K32,N13:N32,O13:O32,P13:P32,E13:E32,Q13:Q32,R13:R32,S13:S32,T13:T32),-AGGREGATE(5,6,I13:I32,J13:J32,K13:K32,N13:N32,O13:O32,P13:P32,E13:E32,Q13:Q32,R13:R32,S13:S32,T13:T32),0.000000001)` |

**Charts**, in points from the top left of the sheet.

| Name | Type | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|---|
| `tier_wf_py` | Stacked Bar | 1177.5 | 86 | 400 | 402 | 3 | -0.03 to 1.0556 |
| `tier_wf_ac` | Stacked Bar | 1587.5 | 86 | 300 | 402 | 3 | -0.03 to 1.0556 |
| `tier_var_abs` | Clustered Bar | 1897.5 | 86 | 175 | 402 | 5 | -0.16 to 0.344 |
| `tier_var_rel` | Clustered Bar | 2082.5 | 86 | 200 | 402 | 5 | -35 to 265 |

| Chart | Series | Fill | Marker | Stem | Labelled |
|---|---|---|---|---|---|
| `tier_wf_py` | `PY base` | no fill | None | - |  |
| `tier_wf_py` | `PY adds` | solid #7F7F7F | None | - | yes |
| `tier_wf_py` | `PY subs` | solid #BFBFBF | None | - | yes |
| `tier_wf_ac` | `AC base` | no fill | None | - |  |
| `tier_wf_ac` | `AC adds` | solid #404040 | None | - | yes |
| `tier_wf_ac` | `AC subs` | solid #7F7F7F | None | - | yes |
| `tier_var_abs` | `ΔPY` | no fill | None | - | yes |
| `tier_var_abs` | `ΔPY pos good` | solid #8CB400 | None | - |  |
| `tier_var_abs` | `ΔPY pos bad` | solid #FF0000 | None | - |  |
| `tier_var_abs` | `ΔPY neg good` | solid #8CB400 | None | - |  |
| `tier_var_abs` | `ΔPY neg bad` | solid #FF0000 | None | - |  |
| `tier_var_rel` | `ΔPY%` | no fill | None | - | yes |
| `tier_var_rel` | `ΔPY% pos good` | solid #8CB400 | None | - |  |
| `tier_var_rel` | `ΔPY% pos bad` | solid #FF0000 | None | - |  |
| `tier_var_rel` | `ΔPY% neg good` | solid #8CB400 | None | - |  |
| `tier_var_rel` | `ΔPY% neg bad` | solid #FF0000 | None | - |  |

### C13D

> Every month of 2025 sits above the fifteen-year average, Mar by +179%; the book has grown 5.5x since 2011, so the grid reads as a climb from below the line to well above it

**A panel grid.**

| | |
|---|---|
| Cells used | `A1:BU106` |
| Print area | `A109:N148` |
| Formula cells | 2092, in 97 shapes |
| Typed cells | 366 |

**Typed values** - 366 cells in `A11:BN106`. Everything else is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in that range, so go by the shading.

**Every formula on the sheet.** Cells that say the same thing about their own position are one row here, with the range they cover and the first of them written out.

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
| 22 cells in `BB55:BH100` | - | `BB55` | `=NA()` |
| `B26:M26` | average 15 years | `B26` | `=AVERAGE(B11:B25)` |
| `B35:M35` | Jan, Feb, Mar, Apr | `B35` | `=(B11-B26)/B26*100` |
| `B36:M36` | 2012 | `B36` | `=(B12-B26)/B26*100` |
| `B37:M37` | 2013 | `B37` | `=(B13-B26)/B26*100` |
| `B38:M38` | 2014 | `B38` | `=(B14-B26)/B26*100` |
| `B39:M39` | 2015 | `B39` | `=(B15-B26)/B26*100` |

*29 further named shapes are not listed - this sheet has 69 in all.*

28 unnamed one-off formulas: `BL53`, `BO53`, `BL54`, `BO54`, `BL55`, `BO55`, `BL56`, `BO56`, `BL57`, `BO57`, `BL58`, `BO58`, ....

**Charts**, in points from the top left of the sheet.

| Name | Type | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|---|
| `panel_grid` | chart type -4111 | 0 | 1646.5 | 720 | 500 | 13 | -0.09 to 4 |
| `reference_panel` | Clustered Column | 549.9 | 2025.3 | 155.1 | 95.9 | 1 | 0 to 5184 |

| Chart | Series | Fill | Marker | Stem | Labelled |
|---|---|---|---|---|---|
| `panel_grid` | `ΔØ %` | head #404040, border #404040 | Square | 1.2pt #8CB400 |  |
| `panel_grid` | ` ` | head #404040, border #404040 | Square | 1.2pt #FF0000 |  |
| `panel_grid` | `ΔØ %` | head #404040, border #404040 | Square | 1.2pt #8CB400 |  |
| `panel_grid` | ` ` | head #404040, border #404040 | Square | 1.2pt #FF0000 |  |
| `panel_grid` | `ΔØ %` | head #404040, border #404040 | Square | 1.2pt #8CB400 |  |
| `panel_grid` | ` ` | head #404040, border #404040 | Square | 1.2pt #FF0000 |  |
| `panel_grid` | `ΔØ %` | head #404040, border #404040 | Square | 1.2pt #8CB400 |  |
| `panel_grid` | ` ` | head #404040, border #404040 | Square | 1.2pt #FF0000 |  |
| `panel_grid` | `divider` | solid #000000 | None | - |  |
| `panel_grid` | `bandrule` | solid #000000 | None | - |  |
| `panel_grid` | `baseline` | solid #000000 | None | - |  |
| `panel_grid` | `ptitle` | solid #000000 | None | - | yes |
| `panel_grid` | `ytick` | solid #000000 | None | - |  |
| `reference_panel` | `Series1` | solid #156082 | None | - | yes |

### T01B

> Every columned segment beat the 96 combined ratio target in the year to date - companywide +6,536 $m against the +1,996 the target implies. Agency contributes the largest excess (+2,037 $m) on 27% less earned premium than Direct

**A table with drawn columns.**

| | |
|---|---|
| Cells used | `A1:O23` |
| Print area | `A9:O23` |
| Formula cells | 70, in 12 shapes |
| Typed cells | 30 |

**Typed values** - 30 cells in `A15:K20`. Everything else is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in that range, so go by the shading.

**Every formula on the sheet.** Cells that say the same thing about their own position are one row here, with the range they cover and the first of them written out.

| Cells | Named | Example | Formula |
|---|---|---|---|
| 14 cells in `D15:L21` | ΔPY | `D15` | `=C15-A15` |
| 14 cells in `F15:N21` | ΔPL | `F15` | `=C15-B15` |
| 12 cells in `E15:M21` | ΔPY% | `E15` | `=IF(A15<=0,NA(),(C15-A15)/A15*100)` |
| 12 cells in `G15:O21` | ΔPL% | `G15` | `=IF(B15<=0,NA(),(C15-B15)/B15*100)` |
| 6 cells in `A18:K18` | - | `A18` | `=ABS(SUM(A15:A17))` |
| 6 cells in `A21:K21` | - | `A21` | `=SUM(A15:A17)+SUM(A19:A20)` |
| `B7` | Subject | `B7` | `=B2&IF(B3="",""," in "&B3)` |

5 unnamed one-off formulas: `A9`, `A10`, `A11`, `A13`, `I13`.

### T02A

> Against the 96 combined ratio target, every columned segment is ahead for the year to date - +6,536 $m against +1,996. The bars are the gap to target, drawn at one scale for the month and the year so the two can be read against each other

**A table with drawn columns.**

| | |
|---|---|
| Cells used | `A1:AD23` |
| Print area | `A9:L23` |
| Formula cells | 158, in 27 shapes |
| Typed cells | 30 |

**Typed values** - 30 cells in `A15:I20`. Everything else is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in that range, so go by the shading.

**Every formula on the sheet.** Cells that say the same thing about their own position are one row here, with the range they cover and the first of them written out.

| Cells | Named | Example | Formula |
|---|---|---|---|
| 14 cells in `D15:J21` | ΔPL | `D15` | `=C15-B15` |
| 12 cells in `E15:K21` | ΔPL% | `E15` | `=IF(B15<=0,NA(),(C15-B15)/B15*100)` |
| `O15:O21` | ΔPL scaled | `O15` | `=IF(ISBLANK(D15),NA(),D15/$M$15)` |
| `P15:P21` | ΔPL text | `P15` | `=IF(ISBLANK(D15),"",TEXT(D15,"[>0.5]+0;[<-0.5]-0;0"))` |
| `Q15:Q21` | ΔPL good | `Q15` | `=IF(ISBLANK(D15),NA(),IF(D15>0,D15/$M$15,NA()))` |
| `R15:R21` | ΔPL bad | `R15` | `=IF(ISBLANK(D15),NA(),IF(D15<0,D15/$M$15,NA()))` |
| `S15:S21` | ΔPL% scaled | `S15` | `=IF(ISBLANK(E15),NA(),E15/$N$15)` |
| `T15:T21` | ΔPL% text | `T15` | `=IF(ISBLANK(E15),"",TEXT(E15,"[>0.05]+0.0""%"";[<-0.05]-0.0""%"";0.0""%"""))` |
| `U15:U21` | ΔPL% good | `U15` | `=IF(ISBLANK(E15),NA(),IF(E15>0,E15/$N$15,NA()))` |
| `V15:V21` | ΔPL% bad | `V15` | `=IF(ISBLANK(E15),NA(),IF(E15<0,E15/$N$15,NA()))` |
| `W15:W21` | ΔPL scaled | `W15` | `=IF(ISBLANK(J15),NA(),J15/$M$15)` |
| `X15:X21` | ΔPL text | `X15` | `=IF(ISBLANK(J15),"",TEXT(J15,"[>0.5]+0;[<-0.5]-0;0"))` |
| `Y15:Y21` | ΔPL good | `Y15` | `=IF(ISBLANK(J15),NA(),IF(J15>0,J15/$M$15,NA()))` |
| `Z15:Z21` | ΔPL bad | `Z15` | `=IF(ISBLANK(J15),NA(),IF(J15<0,J15/$M$15,NA()))` |
| `AA15:AA21` | ΔPL% scaled | `AA15` | `=IF(ISBLANK(K15),NA(),K15/$N$15)` |
| `AB15:AB21` | ΔPL% text | `AB15` | `=IF(ISBLANK(K15),"",TEXT(K15,"[>0.05]+0.0""%"";[<-0.05]-0.0""%"";0.0""%"""))` |
| `AC15:AC21` | ΔPL% good | `AC15` | `=IF(ISBLANK(K15),NA(),IF(K15>0,K15/$N$15,NA()))` |
| `AD15:AD21` | ΔPL% bad | `AD15` | `=IF(ISBLANK(K15),NA(),IF(K15<0,K15/$N$15,NA()))` |
| 6 cells in `A18:I18` | - | `A18` | `=ABS(SUM(A15:A17))` |
| 6 cells in `A21:I21` | - | `A21` | `=SUM(A15:A17)+SUM(A19:A20)` |
| `M15:N15` | span unit, span rel | `M15` | `=MAX(AGGREGATE(4,6,D15:D21,J15:J21),-AGGREGATE(5,6,D15:D21,J15:J21),0.000000001)` |
| `B7` | Subject | `B7` | `=B2&IF(B3="",""," in "&B3)` |

5 unnamed one-off formulas: `A9`, `A10`, `A11`, `A13`, `G13`.

**Charts**, in points from the top left of the sheet.

| Name | Type | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|---|
| `panel_dpl_month` | Clustered Bar | 170.4 | 640.5 | 51 | 108.8 | 3 | -0.0348 to 0.4434 |
| `panel_dplp_month` | Clustered Bar | 217.5 | 640.5 | 96.7 | 108.8 | 3 | -0.0345 to 0.9321 |
| `panel_dpl_ytd` | Clustered Bar | 577 | 640.5 | 106.2 | 108.8 | 3 | -0.0435 to 1.0347 |
| `panel_dplp_ytd` | Clustered Bar | 679.5 | 640.5 | 106.2 | 108.8 | 3 | -0.0345 to 1.0356 |

| Chart | Series | Fill | Marker | Stem | Labelled |
|---|---|---|---|---|---|
| `panel_dpl_month` | `ΔPL` | no fill | None | - | yes |
| `panel_dpl_month` | `ΔPL good` | solid #8CB400 | None | - |  |
| `panel_dpl_month` | `ΔPL bad` | solid #FF0000 | None | - |  |
| `panel_dplp_month` | `ΔPL%` | no fill | None | - | yes |
| `panel_dplp_month` | `ΔPL% good` | solid #8CB400 | None | - |  |
| `panel_dplp_month` | `ΔPL% bad` | solid #FF0000 | None | - |  |
| `panel_dpl_ytd` | `ΔPL` | no fill | None | - | yes |
| `panel_dpl_ytd` | `ΔPL good` | solid #8CB400 | None | - |  |
| `panel_dpl_ytd` | `ΔPL bad` | solid #FF0000 | None | - |  |
| `panel_dplp_ytd` | `ΔPL%` | no fill | None | - | yes |
| `panel_dplp_ytd` | `ΔPL% good` | solid #8CB400 | None | - |  |
| `panel_dplp_ytd` | `ΔPL% bad` | solid #FF0000 | None | - |  |

### T03A

> A 87.4 combined ratio against the 96 target left an underwriting result of +10,289 $m, +7,023 more than the target implies. The gain over 2024 is all in losses: the loss ratio fell 3.2 points while the expense ratio rose 1.8

**A table with drawn columns.**

| | |
|---|---|
| Cells used | `A1:H24` |
| Print area | `A9:H24` |
| Formula cells | 48, in 12 shapes |
| Typed cells | 12 |

**Typed values** - `B15:D18`. Everything else is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in that range, so go by the shading.

**Every formula on the sheet.** Cells that say the same thing about their own position are one row here, with the range they cover and the first of them written out.

| Cells | Named | Example | Formula |
|---|---|---|---|
| `E15:E22` | ΔPY | `E15` | `=D15-B15` |
| `F15:F22` | ΔPY% | `F15` | `=IF(B15<=0,NA(),(D15-B15)/B15*100)` |
| `G15:G22` | ΔPL | `G15` | `=D15-C15` |
| `H15:H22` | ΔPL% | `H15` | `=IF(C15<=0,NA(),(D15-C15)/C15*100)` |
| `B19:D19` | Underwriting expenses | `B19` | `=ABS(-SUM(B16:B18))` |
| `B20:D20` | Underwriting result | `B20` | `=B15-SUM(B16:B18)` |
| `B21:D21` | Loss and LAE ratio | `B21` | `=IF(B15=0,NA(),B16/B15*100)` |
| `B22:D22` | Combined ratio | `B22` | `=IF(B15=0,NA(),B19/B15*100)` |
| `B7` | Subject | `B7` | `=B2&IF(B3="",""," in "&B3)` |

3 unnamed one-off formulas: `A9`, `A10`, `A11`.

### T04A

> A 87.4 combined ratio against the 96 target left an underwriting result of +10,289 $m, +7,023 more than the target implies. The gain over 2024 is all in losses: the loss ratio fell 3.2 points while the expense ratio rose 1.8

**A table with drawn columns.**

| | |
|---|---|
| Cells used | `A1:Q24` |
| Print area | `A9:G24` |
| Formula cells | 98, in 23 shapes |
| Typed cells | 12 |

**Typed values** - `B15:D18`. Everything else is a formula over these. The cells a reader is meant to edit are shaded; a sheet may also park a constant the engine needs in that range, so go by the shading.

**Every formula on the sheet.** Cells that say the same thing about their own position are one row here, with the range they cover and the first of them written out.

| Cells | Named | Example | Formula |
|---|---|---|---|
| `E15:E22` | ΔPL | `E15` | `=D15-C15` |
| `F15:F22` | ΔPL% | `F15` | `=IF(C15<=0,NA(),(D15-C15)/C15*100)` |
| `J15:J22` | ΔPL scaled | `J15` | `=IF(ISBLANK(E15),NA(),E15/$H$15)` |
| `K15:K22` | ΔPL text | `K15` | `=IF(ISBLANK(E15),"",TEXT(E15,"[>0.5]+0;[<-0.5]-0;0"))` |
| `N15:N22` | ΔPL% scaled | `N15` | `=IF(ISBLANK(F15),NA(),F15/$I$15)` |
| `O15:O22` | ΔPL% text | `O15` | `=IF(ISBLANK(F15),"",TEXT(F15,"[>0.05]+0.0""%"";[<-0.05]-0.0""%"";0.0""%"""))` |
| 4 cells in `L15:L22` | ΔPL good | `L15` | `=IF(ISBLANK(E15),NA(),IF(E15>0,E15/$H$15,NA()))` |
| 4 cells in `M15:M22` | ΔPL bad | `M15` | `=IF(ISBLANK(E15),NA(),IF(E15<0,E15/$H$15,NA()))` |
| 4 cells in `P15:P22` | ΔPL% good | `P15` | `=IF(ISBLANK(F15),NA(),IF(F15>0,F15/$I$15,NA()))` |
| 4 cells in `Q15:Q22` | ΔPL% bad | `Q15` | `=IF(ISBLANK(F15),NA(),IF(F15<0,F15/$I$15,NA()))` |
| `L16:L19` | Losses and loss adjustment expense | `L16` | `=IF(ISBLANK(E16),NA(),IF(E16<0,E16/$H$15,NA()))` |
| `M16:M19` | Losses and loss adjustment expense | `M16` | `=IF(ISBLANK(E16),NA(),IF(E16>0,E16/$H$15,NA()))` |
| `P16:P19` | Losses and loss adjustment expense | `P16` | `=IF(ISBLANK(F16),NA(),IF(F16<0,F16/$I$15,NA()))` |
| `Q16:Q19` | Losses and loss adjustment expense | `Q16` | `=IF(ISBLANK(F16),NA(),IF(F16>0,F16/$I$15,NA()))` |
| `B19:D19` | Underwriting expenses | `B19` | `=ABS(-SUM(B16:B18))` |
| `B20:D20` | Underwriting result | `B20` | `=B15-SUM(B16:B18)` |
| `B21:D21` | Loss and LAE ratio | `B21` | `=IF(B15=0,NA(),B16/B15*100)` |
| `B22:D22` | Combined ratio | `B22` | `=IF(B15=0,NA(),B19/B15*100)` |
| `H15:I15` | span unit, span rel | `H15` | `=MAX(AGGREGATE(4,6,E15:E22),-AGGREGATE(5,6,E15:E22),0.000000001)` |
| `B7` | Subject | `B7` | `=B2&IF(B3="",""," in "&B3)` |

3 unnamed one-off formulas: `A9`, `A10`, `A11`.

**Charts**, in points from the top left of the sheet.

| Name | Type | Left | Top | Width | Height | Series | Axis |
|---|---|---|---|---|---|---|---|
| `panel_dpl` | Clustered Bar | 375.5 | 538.9 | 106.2 | 123.8 | 3 | -1.0849 to 1.0849 |
| `panel_dplp` | Clustered Bar | 478 | 538.9 | 106.2 | 123.8 | 3 | -0.1029 to 1.0401 |

| Chart | Series | Fill | Marker | Stem | Labelled |
|---|---|---|---|---|---|
| `panel_dpl` | `ΔPL` | no fill | None | - | yes |
| `panel_dpl` | `ΔPL good` | solid #8CB400 | None | - |  |
| `panel_dpl` | `ΔPL bad` | solid #FF0000 | None | - |  |
| `panel_dplp` | `ΔPL%` | no fill | None | - | yes |
| `panel_dplp` | `ΔPL% good` | solid #8CB400 | None | - |  |
| `panel_dplp` | `ΔPL% bad` | solid #FF0000 | None | - |  |

---

*Generated by `ibcs_doc.py` from the workbook itself. Every formula, colour and measurement above was read back out of it.*
