# Deploy These Macro Updates to SolidWorks

## Optional dual-entry-point UI

The five existing macro sources are unchanged by this addition. The actual
manual entry point in the supplied source is `RunPackAndGo()` in module
`packngo1` (not a procedure named `packngo()`). Keep any existing manual
button/shortcut pointing at that procedure.

### Install into the existing Pack-and-Go VBA project

1. In SolidWorks, edit `packngo.swp`. Its original standard module must
   remain named `packngo1`.
2. Use **File > Import File** to import
   [userform_automation.bas](userform_automation.bas). If pasting instead,
   create a standard module named `userform_automation` and paste everything
   after the `Attribute VB_Name` line.
3. Also use **File > Import File** to import
   [UIDesignTableSession.cls](UIDesignTableSession.cls). This is a class
   module, not a standard module. It updates the embedded design table and
   leaves the original unsaved for discard; no Excel VBA-project access is needed.
   Also import [UIReferenceSaveGuard.cls](UIReferenceSaveGuard.cls), which
   protects loaded template references against suppression-time save prompts.
   Import [UIRunLogger.cls](UIRunLogger.cls) for local timing diagnostics.
4. Use **Insert > UserForm** and set its **(Name)** to
   `UserForm_AutomationUI`. Open its code window and paste the **entire**
   [UserForm_AutomationUI.frm](UserForm_AutomationUI.frm) source.
   This file is **code-behind, not a native exported form package**: do not
   use File > Import File on it. All controls and their event handlers are
   created/wired at runtime, so no manual layout or `.frx` is needed.
5. Import [UIProgressSession.cls](UIProgressSession.cls) as another class
   module. Insert a second UserForm named `UserForm_AutomationProgress` and
   paste the entire [UserForm_AutomationProgress.frm](UserForm_AutomationProgress.frm)
   code into its code window. Like the input form, this is code-behind,
   not a directly importable form package.
   For standalone quotes, import [UIQuoteResult.cls](UIQuoteResult.cls) as a
   class and [ui_quote_calculator.bas](ui_quote_calculator.bas) as a standard
   module. Insert a UserForm named `UserForm_QuoteReport` and paste the entire
   [UserForm_QuoteReport.frm](UserForm_QuoteReport.frm) code behind it.
   Import [ui_quote_nesting.bas](ui_quote_nesting.bas) as a standard module
   for the sheet-layout estimate.
   Import [UIQuoteComponent.cls](UIQuoteComponent.cls) as a class for the
   component area breakdown. For an existing quote installation, replace
   `UIQuoteResult` and `ui_quote_calculator` and add `UIQuoteComponent`.
6. Ensure references to **SOLIDWORKS type library**, **SOLIDWORKS Constant
   type library**, and **Microsoft Forms 2.0 Object Library** are enabled.
   Inserting a UserForm normally adds the Forms reference automatically.
7. The supplied Pack-and-Go source references undeclared
   `filteredExcludedFiles` and `excludedFiles`. The new module supplies
   those declarations without changing the original file. If your live
   project already declares them in another module, omit the two new
   compatibility declarations to avoid duplicate public names.
8. Use **Debug > Compile**, then save `packngo.swp`. Assign a second macro
   button to module `userform_automation`, procedure `RunWithUI`.
   Do **not** import the four export modules into this project: their
   duplicate public names would conflict. They remain standalone `.swp`
   projects in `packngo1.MACRO_FOLDER`, as before.

**Updating an already-installed UI:** replace the code in the existing
`userform_automation` module (omit its `Attribute VB_Name` line when pasting),
replace the entire input form's code-behind, and install the design-table
class and progress class/form above. Do not
import a duplicate standard module. Recompile and save. The old equation-write
version is not suitable for this design-table-controlled template.

### Progress window and estimates

After clicking Run, a modeless progress window shows the current stage,
completed milestone count/bar, elapsed time at the last update, and an
estimated remaining time when historical timings exist for every remaining
stage. Model-only runs have nine stages; selected exports add one stage each.
Preview-only has two stages: table update/rebuild, then model verification
and releasing temporary reference protection.
Percent complete counts milestones equally, **not files or time**.

The first run displays an unavailable estimate. Successful workflow runs
record each stage's elapsed seconds in the current user's VBA settings:
`HKEY_CURRENT_USER\Software\VB and VBA Program Settings\PlanterWallAutomation\StageSecondsV1`.
Future estimates sum the last recorded durations for the remaining selected
stages. These are rough estimates, not guarantees: assembly size, network
speed and exporter dialog waiting time can change them. Export timings measure
macro invocation duration, not confirmed export success.

VBA repaints and processes events at stage boundaries. A synchronous
SolidWorks/Excel/export macro call can block VBA for minutes; the stage label
remains the last reported activity and the timer/bar may not refresh until
that call returns. There is no background countdown, per-file export progress
or cancellation button. The unchanged exporters do not expose that progress.
The window closes on completion or failure; failures release temporary
settings and show the existing error dialogs. Timing-storage errors are
reported separately without claiming the model workflow failed.

**Wait for the current run to finish before updating/recompiling the macro.**
These additions cannot attach a progress window to an already-running macro.
While a modeless progress window is displayed, do not change the active
document/configuration or start another macro; repaint/event processing is
for visibility, not permission to edit the model mid-workflow.

### Preview-only mode

Select **Preview only (no files created)** in `RunWithUI` to update the
embedded design table and rebuild/verify the template without Pack-and-Go,
output-folder creation or any exporter. Quote-only and all export checkboxes
are cleared and disabled. The destination field/browser are disabled and
no existing output folder is required. Leaving preview mode reenables the
controls but does not restore previous export selections; select them again.

The model stays open with **unsaved** preview changes and is zoomed to fit.
Do not save the template: close it without saving and reopen a clean template
before another preview or full run. Existing clean-template preflight remains
in force; this version does not permit reuse of a dirty preview. Close any
other affected template documents without saving as well. Preview restores
temporary reference read-only states and application preferences on completion
or failure; no original design-table restoration is attempted.

The preview still writes a local timing log (`RUN_PREVIEW` on completion).
"No files created" means no model/export deliverables, not absence of logs.
To install this UI change, replace `userform_automation` and the entire code
behind the existing `UserForm_AutomationUI`, compile and save. No new class
or form is required.

### Standalone quote-only quantities

Quote-only now uses only the validated UI inputs and VBA arithmetic. It does
not access an active document, equation manager, design table, Excel or runtime
logger. No clean template, destination folder or exporters are needed; no files
are written. Launch `RunWithUI` even with no document open, choose **Quote only
(input-only material estimate)** and click Run. The results window opens with
a header showing the material-only cost and selected design, summary tiles,
a wall/base/end component table and a separate gusset table.
**Technical Details** switches to the
full scrollable, copyable calculation report; **Back to Summary** returns.
It is **not a final manufacturing quote**.
No model changes need discarding after a quote.

All export checkboxes start unchecked. Quote-only clears/disables every export,
preview and the output folder/browser. Leaving quote mode reenables the controls
without restoring old selections. Preview and full-model modes still require
the original clean template. Model initialization/validation now occurs after
form acceptance, only for those modes. The form no longer reads template
equations to prefill choices: it uses existing UI globals when available,
otherwise one wall, no return, Mild Steel and 3/16-inch thickness. Dimensions
must be entered explicitly.

Formula version 3 reproduces the supplied workbook N10/N11 and exported
assembly equations:

- Center available length is L for one wall, L - H/2 for two, L - H for
  three/four. Sizing segments = ceiling(available/120); center length =
  available/sizing segments.
- Side sizing segments = ceiling(W/120). Side length = W/segments, plus
  thickness + thickness*(segments-2)/segments when segments > 1.
- Assembly center count = Int((available-1)/120)+1; assembly side count =
  Int((W-1)/120)+1. These are **per wall run**, not overall totals.
- Gussets per planter: 1 up to 24 inches, 2 up to 48, 3 up to 72, 4 up to 96,
  otherwise 5, matching the source's inclusive thresholds.
- One/two/three/four wall layouts infer respectively 1/1/1/2 center wall runs
  and 0/1/2/2 side wall runs from the top-level suppression rules. Totals
  multiply per-run counts; they must be checked against actual known designs.
  Left/Right affects the
  two-wall orientation, not these aggregate quantities.

The source sizing/count rules disagree for fractional dimensions just above
120-inch multiples (e.g. 120.5: ceiling gives 2, assembly count gives 1).
The report preserves both rules, warns, and withholds aggregate totals for
affected layouts. Unused side counts in one-wall layouts do not invalidate
the center totals. Do not silently substitute a new count rule.

Reference example: four walls, 300 x 450 x 48 inches, 1/4-inch Mild Steel,
double return gives center length 84, side length 112.875, center pattern
count 3, side count 4, center gussets 4, side gussets 5, mates 66 and 234.
Totals: 6 center planters and 8 side planters. Following the supplied component
specification, each planter has two mounting gussets and its gusset equation
value minus one internal center gussets. This gives 28 mounting gussets and
50 internal center gussets (78 total), replacing the earlier equation-value
totals of 24 and 40, which were not physical component quantities.

The component report implements [Component_Quote_Formulas.txt](../Component_Quote_Formulas.txt):

- Eight groups: center wall/base, side wall/base, left/right ends, mounting
  gussets and internal center gussets. Side wall variants share the supplied
  dimensions and are aggregated rather than double-counted.
- Side wall/base length = side planter length - 2*thickness, confirmed by the
  user (112.875 - 0.5 = 112.375 inches in the reference example).
- Left/right ends each occur once per side run, not once per side segment.
- Walls use calibrated flat heights. For 1/4-inch material, single-return
  offset is 0.5956 and double-return offset 1.1912 inches; for 3/16-inch,
  0.6798 and 1.3595. A 2-inch single return adds one more inch. No return
  adds nothing. Preserve the measured rounded double offset independently.
  Measurements were supplied at H=48; other heights remain predictions,
  including H=23.75. These are calibrated rules, not exact theoretical bend
  calculations or live SolidWorks measurements.
- Gusset height = H-2.1-T; bottom width = H/2-T-0.5; top width = bottom/3.
  Area = (top+bottom)/2*height; both gusset types use that same trapezoid.
- Each row reports material, thickness, governing rules, dimensions, profile
  area, quantity and total area. Count mismatches withhold quantities and
  aggregate areas; per-part dimensions/areas remain visible.
- The reference double-return example totals 162600.4588 in^2 before any
  cutout deduction. This is summed profile area, not purchased sheet area.

Holes, notches and reliefs are **not deducted**: the supplied rectangular/
trapezoidal profiles are explicitly labeled. The sheet estimator below uses
120 x 60-inch stock. Profile (part) material costs above do not include
offcuts; the quote pricing below charges full purchased sheets.
The local workbook and `equations.txt` export remain Git-ignored.

#### Profile weight and material pricing

The quote-only **Quote material ($/lb)** input defaults to **0.90**, can be
changed for each run under **Properties...**, and is ignored for preview and
full model runs.
It accepts positive dot-decimal numbers, without currency symbols, commas,
fractions or units. Zero/negative/missing prices produce explicit errors.
The API also validates positive prices; its optional default is 0.90.

Both Mild Steel and Borcon Weathering Steel use the user-approved estimating
density **7850 kg/m^3**, converted as `7850 * 0.0254^3 / 0.45359237` =
approximately **0.2835992422 lb/in^3**. This is a typical carbon-steel value,
not a verified BORCON-specific material property:

- [Carbon-steel density reference](https://lyhsteel.com/steel-calculators/metal-density-chart/)
  lists 7850 kg/m^3 as an estimating value.
- [Border Concepts BORCON product description](https://www.borderconcepts.com/products/border-line-steel-edging-borcon-borcon-l416)
  identifies weathering steel but does not specify density or an ASTM grade.

Weight = profile area * thickness * density; cost = weight * entered $/lb.
Per-part and quantity-extended weights/costs are reported. Aggregate results
are withheld on count mismatches. Calculations retain precision; displayed
costs are formatted to two decimal places, and totals are calculated from
unrounded values rather than summing rounded row displays.

These are undeducted profile-material estimates, not true net solid mass.
They exclude sheet offcuts, nesting losses, cutting/bending/welding/finishing,
labor and margin. No density difference between material labels is invented;
different purchase prices can be entered for each quote.
For the 400 x 400 x 23.75-inch, four-wall, 3/16-inch, single 1-inch return
example, profile weight is approximately **3806.4128 lb**, and material cost
at 0.90/lb is **$3425.77**.

Updating an existing installation: replace `UserForm_AutomationUI`,
`ui_quote_calculator`, `UIQuoteResult` and `UIQuoteComponent`, then compile
and save.

#### Labor, paint, margin and sell price

Quote settings are grouped under the **Properties...** button on the input
form (an overlay panel with **Restore Defaults** and **Done**). They are used
only by Quote only:

| Property | Default | Applies to |
|---|---|---|
| Quote material ($/lb) | 0.90 | Full purchased sheets: sheets x 120 x 60 x T x density |
| Cutting labor ($/sheet) | 75.00 | Every purchased sheet |
| Bending labor ($/bent part) | 5.00 | Center/side planter walls and left/right side walls, only when a return is selected |
| Welding labor ($/part) | 10.00 | Every physical part (all parts are welded) |
| Paint ($/sq ft, both sides) | 2.50 | 2 x total profile area, only when **Painted** is on |
| Default margin (%) | 50.00 | Gross margin, 0-95% |
| Quote part gap / edge margin (in) | 0.25 / 0.25 | Sheet nesting |

Labor and paint defaults are placeholders until time studies are available.
The **Painted / Not painted** toggle sits next to Material.

Total cost = purchased-sheet material + labor + paint. Sell price uses gross
margin: `sell = total cost / (1 - margin)`, so the default 50% margin means
sell price is twice total cost. Pricing is withheld when quantities disagree or no
sheet layout is available.

The results page leads with the sell price and shows a pricing row (material,
labor, paint, total cost, margin). The margin box and spinner recalculate the
sell price and profit immediately; hover any pricing tile for its breakdown.
Technical Details includes the full pricing calculation.

#### Field assembly estimates

In **Quote only**, **Quantity** defaults to 1 and accepts whole numbers from
1 to 2000. It is disabled for model/preview runs and does not duplicate CAD
assemblies. Quantity scales component counts, profile totals, welding and
paint; all copies of that design are re-nested together, so purchased sheets
and cutting labor are recalculated rather than multiplying rounded prices.
The existing limit of 2000 physical components **per design including copies**
still applies; an over-limit design has no layout or valid quote price.

All quote-only runs use the field assembly, even for a single design.
**Run field** is disabled until at least one planter is added.
Quantity and design name are in the right-hand **FIELD ASSEMBLY - QUOTE ONLY**
panel, not the main design inputs. **Preview only** sits to the right of
**Quote only**; these remain mutually exclusive input modes.

1. Enter a design, quantity and optional design name in the right-hand panel.
2. Select material, thickness, Painted and any quote Properties for that design.
3. Click **Add planter**. This captures a snapshot of the current inputs and rates.
4. Select a saved row to load its name, quantity, dimensions, wall/return type,
   material, thickness, Painted choice and quote properties. Make changes and
   click **Save changes** to recalculate and replace that row in place.
   **New design** deselects the row and resets name/quantity, retaining other
   inputs as a starting point; **Add planter** always adds another row.
5. Click **Run field**. **Only saved rows are quoted**; draft or unsaved input
   edits are not included. Remove all rows to leave Quote only mode.

Each design nests on its own sheets, even when materials match, to preserve
section separation for Operations. Designs may use different materials,
thicknesses, paint choices and prices without sharing stock.
The field results list shows each design's quantity, sheets, total cost and
sell price, plus combined material, labor, paint, cost, profit and sell price.
The sell price is the primary headline, with profit/total cost as secondary
cards and material/labor/paint/sheets as smaller supporting cards.
No complete field totals are displayed if any design cannot be priced.
**View selected** opens its existing detailed results and layout PDF, with
component quantities and layouts covering **all copies of that design**.
The PDF identifies the design and quantity; a combined field PDF is not generated.
The design/quantity line sits above the PDF summary cards with a reserved gap.
Field-summary column headings use the same widths as their data columns.
The margin in Properties sets the **whole field's** margin when running.
The results margin text box and arrow spinner update **every design**
and the field totals immediately (0-95%). Per-design detail margins are
read-only so the field always has one common gross margin. Invalid margin
input is shown in red with a last-valid-price notice.

**Preview selected** hides the field results while running the same guarded
preview workflow as the input UI for **one copy** of the selected design.
The original, clean `planter_assembly.SLDASM` must be active. Preview updates
the unsaved template only; it creates no Pack-and-Go copies or exports, and
does not modify quote quantities or prices. The field results return after the
preview; close the results to inspect the model. Close the template **without
saving** and reopen it before another preview or model run. Preview failures
are reported without discarding the field quote.
No per-copy sheet/cost dropdown is shown, because copies share a nest.

Install/update: replace `userform_automation`, `ui_quote_calculator`,
`ui_quote_layout_pdf`, `UIQuoteResult`, and the code behind
`UserForm_AutomationUI` and `UserForm_QuoteReport`. Import the new `ui_quote_field.bas` module and create
a new UserForm named **UserForm_FieldQuote**, pasting its `.frm` code starting
at `Option Explicit` (not a designer-form import). Keep the existing
`UserForm_QuoteReport`, then Debug > Compile and save.

#### Sheet nesting estimate and Operations verification

Quote-only now includes editable **Quote part gap (in)** and **Quote edge
margin (in)** inputs under **Properties...**, both defaulting to **0.25**.
They are ignored for model and preview runs. Margins apply at all four sheet edges. Rotation
by 90 degrees is allowed, as approved by the user. The current quote contains
one material/thickness; sheets are not shared between separate quotes.

Operations confirmed **0.25 in is a good part-gap baseline**, with typical
shop settings of **0-0.25 in**. Keep 0.25 as the default; zero is an optional
touching-edge estimate for flat panels. Zero gap does not generate shared
cut paths, reduce cutting labor automatically, or establish production
clearances. Operations sets up "co-edge" / common-line cutting in CypCut.
Other nonnegative gaps remain available for job-specific requirements.

The workflow remains: **individual component DXFs -> quantities in CypCut ->
Operations-controlled nesting -> cutting**. The PDF estimates purchased stock
and utilization; it is not a production nest or toolpath. Operations may group
large jobs by section to simplify unloading, which can change sheet usage.
No combined nest DXF or section-specific production nesting is added.

CypCut can rotate in finer increments (Operations mentioned 20 degrees);
the estimator still tests 0/90-degree bundle rotations. Its bounding-box
packing cannot safely gain arbitrary-angle rotations without polygon-level
collision/spacing checks. Finer angles may improve a production nest, while
section grouping may reduce utilization; no fixed adjustment factor is applied.
The PDF states this distinction and flags zero-gap estimates explicitly.

These settings are **estimating assumptions, not an ANSI/ISO/ASTM
laser-spacing standard**. The published
[Lapas spacing guidance](https://lapas.io/blog/part-spacing-nesting/)
describes process/thickness-dependent starting points (e.g. 2.5 mm for fiber
laser and 3.5 mm for CO2 laser at 6 mm thickness), and calls for shop testing.
It does not prescribe our 0.25-inch defaults. Operations has confirmed the
6.35 mm part-gap baseline; the sheet-edge margin remains an assumption, not a value
certified by that source.

**Operations must verify before quoting purchased sheets or using a layout:**

- Machine/process, assist gas, lead-in/pierce clearance and usable sheet size.
- Part spacing, edge trim, clamps and sheet handling constraints.
- Rotation/grain/finish restrictions and whether opposed gusset pairs work.
- Thermal cut sequence, small-part support, tabs and remnant reuse policy.
- Actual cutouts/reliefs and calibrated flat-pattern measurements.

Algorithm:

1. Expand all quantities into physical parts; skip zero quantities.
2. Group same-profile mounting/internal gussets into opposed pairs, allowing
   a mounting/internal mixed pair. Trapezoid height H, bottom B, top A:
   horizontal diagonal separation = `gap * sqrt(1 + ((B-A)/H)^2)`.
   The resulting pair envelope is `(B+A+separation) x H`, containing two
   actual trapezoids with perpendicular clearance equal to the requested
   gap. An odd final gusset uses a single B x H envelope. If the pair envelope
   does not fit the usable sheet in either rotation, fall back to individual
   gussets rather than rejecting otherwise fitting parts.
3. Retain the original 32-trial baseline: four descending part orders (area, longest
   side, shortest side and perimeter), two free-rectangle fit scores (least
   residual area and best short-side fit), and three guillotine split rules.
   The original four layouts run first, and ties keep the earlier result.
   Then try eight MaxRects layouts (four orders x two fit scores). MaxRects
   tracks overlapping candidate free rectangles across earlier split
   boundaries. Every placed envelope is subtracted from all intersecting
   spaces, and contained/duplicate free rectangles are removed so candidates
   never include occupied material. Candidate areas must not be summed as
   remaining stock because they can overlap.
   Then evaluate up to eight all-single-gusset MaxRects layouts and eight
   mixed layouts (approximately half paired and half single), for up to
   48 candidate layouts. This allows singles to fill narrow spaces without
   discarding the paired baseline. Ties retain the earlier candidate.
   If the area lower bound is reached, redundant remaining trials can stop.
   The trial count indicates the configured candidate search, not solver
   iterations. This is heuristic rectangle/envelope nesting, not unrestricted
   polygon nesting or a proven optimum. Pair-envelope wedges remain unused.
4. Respect part gaps and edge margins and allow 0/90-degree bundle rotation.
   A virtual trailing gap avoids incorrectly requiring extra clearance at
   the sheet boundary. Store each bundle's sheet, coordinates, envelope,
   rotation, actual profile area and physical part count for audit.
5. Try to eliminate the least-used sheet by rebuilding MaxRects free spaces
   on the retained sheets and relocating all its bundles. Reserve the same
   padded gap and margins. Commit only a complete, clearance-checked
   elimination; an unsuccessful attempt leaves the source layout unchanged.
   This pass does not rearrange the already placed parts on retained sheets.

Summary reports estimated sheet count, profile utilization and unused/scrap
percentage. Technical Details includes per-sheet utilization, physical part
counts and bundle placements in inches from the sheet's lower-left origin.
The optional PDF preview below visualizes these placements; no cutting program
or machine toolpaths are generated.

`Utilization = total profile area / (sheet count * 7200) * 100`.
`Unused/scrap = 100 - utilization`. The denominator is the full purchased
sheet area, not just the usable area after margins. Pair envelopes are NOT
used as the parts' profile area. Unused area includes potentially reusable
remnants, gaps, margins and shape offcuts; it is not a prediction of discarded
scrap weight. Holes/reliefs remain undeducted, as in the profile calculator.
The area-only lower bound `ceiling(total profile area / 7200)` is reported
separately and is not promised to be a feasible layout.

For 300 x 200 x 20 inches, four walls, 1/4-inch Mild Steel, double return,
the current combined heuristic estimates **7 sheets**, **72.6934% utilization** and
**27.3066% unused area**, versus a 6-sheet area-only lower bound.

Count mismatches, oversized bundles or quotes exceeding 2000 physical parts
produce an explicit unavailable status rather than a misleading sheet count.
Bad/negative settings raise actionable input errors. A blank longer than the
usable length (for example a 120-inch wall with a 0.25-inch margin) is placed
by reducing the edge margin on that axis only, just as far as needed; the
other axis keeps the requested margin. The reduction is never silent: it is
shown in orange under the results sheet table, in Technical Details and on the
PDF summary page, and sheet pages show the effective ends/sides margins.
Parts larger than 120 x 60 even with no margin remain unavailable.
Adjustments require Operations approval. The supported margin is <30 inches.
Zero clearance is supported for exploration, not recommended as production
approval or implicit common-line cutting.

Install on the existing quote version: add `ui_quote_nesting`, replace
`UIQuoteResult`, `ui_quote_calculator`, the input form code and the results
form code, compile and save. No changes to `UIQuoteComponent` are needed.
Live VBA compilation and comparison against the shop's laser nesting
software remain required. Tests validate boundaries, non-overlap/spacing,
rotations, quantities, paired gusset clearance and actual-profile area
conservation, but do not validate the physical cutting process.

Updating from the 32-trial estimator requires replacing
[UIQuoteResult.cls](UIQuoteResult.cls), [ui_quote_nesting.bas](ui_quote_nesting.bas),
[ui_quote_calculator.bas](ui_quote_calculator.bas),
[ui_quote_layout_pdf.bas](ui_quote_layout_pdf.bas), and the code behind
[UserForm_QuoteReport.frm](UserForm_QuoteReport.frm), compiling and saving.
No external dependency or extra optimization button is needed. The search
runs automatically when calculating a quote. Larger component counts take
longer because of repeated packing and free-rectangle pruning.
Tests compare the combined result against the prior 32 baseline layouts over the
quote scenarios, and exercise all eight MaxRects choices for bounds, identity,
area and gap preservation. A regression case reduces four sheets to three
compared with the best of all 24 split trials; this does not imply every
design will improve.

#### Preview sheet layout

**Preview Sheet Layout**, beside Technical Details, generates a PDF in the
Windows temporary folder and opens it in the default PDF viewer. There is no
folder picker or save step. A PDF viewer association is required. Temporary
files use unique `Planter_Layout_*.pdf` names and are never overwritten.
They remain in the temporary folder so an open viewer can continue reading
them; they can be removed after closing the viewer during normal temp cleanup.
Writing/opening errors are reported explicitly. The button is disabled when
the sheet layout is unavailable.

A landscape summary page opens the PDF, explaining flexible gusset trials,
gusset pairing, rotations, clearance and fewest-sheets selection. It shows
total parts, selected sheet count, profile/stock/unused areas, utilization and
the area-only lower bound, selected method and previous baseline sheet count.
This is a best-found estimate, not a guarantee of globally minimum scrap.
Ties keep the first trial; pair-envelope wedges are
not reclaimed, and the area-only bound does not establish a feasible nest.

After the usage table, a Parts List page shows every part used in the layout
(label, description, shape, profile size, total quantity and total area, with a
TOTAL row); parts with zero quantity are omitted.

Following the parts list, one landscape letter-size page per 120 x 60-inch sheet shows proportional
component profiles, the stock border and a dashed usable-area boundary.
Each part is labeled with its PNS component code: CW (Center Planter Wall),
CBP (Center Planter Base Plate), SW (Side Planter Wall), SBP (Side Planter Base
Plate), LW (Left Side Wall), RW (Right Side Wall), MG (Mounting Gusset) and CG
(Center Gusset). Each sheet's legend lists only the parts on that sheet with
their physical quantities. Paired gussets show both trapezoids and diagonal clearance;
rotations and positions match the estimator, including single gussets.
Each page gives material/thickness, gap/margin, sheet and overall utilization
and unused percentages. Unused area includes reusable remnants, not just scrap.

This is an **estimating layout, NOT a production cutting plan**.
It is not full size and has no holes/reliefs, lead-ins, kerf compensation or
cut order. The preview does not open a SolidWorks model, require Excel, or use
the production PDF exporters. PDF generation uses plain VBA and Windows
filesystem/shell services; no additional runtime PDF software is required
besides the viewer.

Install: remove the previous `ui_quote_layout_dxf` module if installed,
import [ui_quote_layout_pdf.bas](ui_quote_layout_pdf.bas), and replace the
code behind `UserForm_QuoteReport`. Compile and save. Test the button with a
valid layout and an unavailable layout, and verify it opens your default viewer.
Re-import [ui_quote_calculator.bas](ui_quote_calculator.bas) too; the
redesigned results header uses its `UIQuoteSummaryCards` function. The original
production DXF exporter is unchanged.

PDF regression tests require `python -m pip install -r tests\requirements.txt`
from the repository root. They load the generated PDF with pypdf in strict
mode and check pages, labels, counts, utilization and profile geometry.
Actual drawn polygons are also checked for overlap and minimum part clearance,
including the diagonal edges of opposed gusset pairs.

#### Results presentation

The summary includes a **Sheet usage - Unused / potential scrap**
table. Each sheet shows physical part quantity (a gusset pair counts as two),
used profile area and unused stock area in square feet, utilization and unused
percentage. A final TOTAL row sums quantities and areas; percentages are
area-weighted, not summed. Each 120 x 60-inch purchased sheet contributes
50 square feet. Unused includes margins, gaps, offcuts and reusable remnants,
not a claim of discarded scrap. Values are displayed to two decimals while
calculations retain full precision. Missing/invalid nesting withholds these
values rather than showing zero unused material.

The results form fits on one screen without form scrolling. A colored header
shows the material estimate (or a review warning) and the design details.
Five tiles show profile weight, profile area, sheet count, utilization and
unused/potential-scrap percentage. The component, gusset and sheet-usage tables sit
below them. Components with a quantity of 0 are hidden; if totals need review,
every row stays visible. Each table is sized to its rows when the window opens,
so there is no blank row at the bottom. Only the sheet-usage list scrolls internally when a design needs
more sheets than fit in its visible rows. Technical Details replaces the
summary area with the full report.

The preview PDF opens with an **Optimization Summary** page. It has
key-metric cards, a used-versus-unused bar, the optimization steps, selected
layout facts and a "Keep in mind" box. Next come paginated sheet-usage tables
with shaded headers, striped rows, right-aligned numbers and a used-versus-unused
bar for each sheet. Totals appear in a highlighted row on the final usage page,
before the individual sheet drawings.

The summary/table rounds displayed dimensions to three decimal inches and
weights/costs to two decimal places. Quantities are whole numbers. The raw
calculation values are not rounded or changed. Technical Details retains the
equation names, per-part profile areas, trapezoid top widths, assumptions and
full-precision calculation values from the earlier report.

Table weights and costs are **quantity-extended totals**, not unit values.
The final **TOTAL MATERIAL - ALL COMPONENTS** row below both tables shows the unrounded calculation's total
weight and cost, formatted to two decimals; both are withheld on count mismatches.
The width/height column is width for base plates and height for walls.
The gusset table has explicit **Height**, **Top width** and **Base width**
columns, plus quantity, total weight and material cost. Gussets are excluded
from the wall/base/end table, but included in the combined total. Omitted components remain
visible with quantity zero. A count mismatch replaces summary totals with
the warning and marks table quantities/weights/costs as Review/Withheld.
Estimate exclusions and calibration/density assumptions remain visible
below either view.

To update only this presentation on an already-installed version 3 quote
calculator, replace `ui_quote_calculator` and the code behind the existing
`UserForm_QuoteReport`, then compile and save. No new class or form is needed.
Test both view transitions, Close/Escape, the reference examples, and a
120.5-inch count-mismatch case. Layout/font fitting needs live MSForms
validation; automated tests do not render the SolidWorks VBA window.

### Runtime logs and performance investigation

Every submitted UI run writes a timestamped UTF-16 tab-separated `.log` in:

```text
C:\Users\lsiewert\StudioProjects\solidworks-planter-wall-automation\run-logs
```

The macro creates this directory automatically. `UI_LOG_FOLDER` in
`userform_automation` is the configurable location: update it for other
machines. The parent repository directory must exist and be writable.
If log creation fails, the run stops before any model changes. If a later
append fails, logging is disabled with an explicit warning while automation
and cleanup continue. Logging has no background process or external service.

Logs include:

- Template path, SolidWorks revision, N3:N9 input values, export selections
  and output location.
- BEGIN/END timing for each progress stage and more detailed Excel open,
  calculation, table commit, reference preflight, template/packed rebuild
  and packed save operations.
- Calculated planter lengths and Pack-and-Go document count.
- SUMMARY duration rows and total runtime.
- Failure details, cleanup duration and INCOMPLETE operations on failed runs.
  Each record is written and closed immediately, so an interrupted run
  retains the events already written even without a final summary.

The final completion/error dialog gives the log path. After a run, leave
the file in `run-logs` and point the assistant at its filename for analysis.
Logs remain **local and Git-ignored**, since they contain engineering inputs
and local/network paths. No log is uploaded or committed automatically.

Timings are wall-clock seconds with subsecond timer resolution, including
time waiting in dialogs. Nested operations overlap: compare stage totals
separately from their child-operation durations rather than summing all rows.
Clock adjustments can affect measurements. Export logs measure the whole
standalone macro invocation; exporter internals are not instrumented here.

Performance candidates to evaluate after collecting a baseline:

1. Duplicate rebuild work: initial template rebuild, post-table rebuild,
   collection rebuild, and packed-copy rebuild. Do not remove them
   merely because there are several; prove geometry/references remain correct.
2. Excel table-open/commit costs versus actual formula calculation time.
3. Pack-and-Go network I/O, drawing inclusion and document count. Compare
   equivalent runs to local and network output folders if stage 6 dominates.
4. Export invocation costs; use the exporters' own logs for individual
   drawing/part detail where available.

Original-table restoration has already been removed on the experimental
branch; stage 7 closes the original without saving. No additional rebuild,
reference inclusion, geometry or exporter behavior is skipped by logging.
Start with a model-only run, then the same inputs with selected exports.
Use another empty output parent for each equivalent run to avoid overwrite
rejection, and note any manual dialog delays when comparing logs.

UI mode now uses a separately instrumented equivalent of manual Step4:
collection rebuild, `GetPackAndGo`, enabling drawing inclusion, and
`GetDocumentNames` each have their own timings. The original manual module
is unchanged. UI mode still rejects failed native collection rather than
using the manual fallback, and it now checks the collection rebuild result.
Discovery prepares one Pack-and-Go object; `SavePackAndGo` is called once
to write the renamed copies. SUMMARY rows repeat measurements, not work.

Before names are replaced, diagnostics record each selected source path,
on-disk file size and metadata lookup duration. Failures are recorded and
warned about without substituting for SolidWorks save-status validation.
This is metadata access, not a full content-read/throughput test, and adds
a separately timed inventory pass. No source file is saved, copied or
changed by that pass. The blocking `SavePackAndGo` call also has its own
timing; the API does not expose internal per-file progress here.

The original closes without saving, so reopening it shows its old saved
inputs by design; the macro does not rewrite those values. The packed model's
equation values are checked against the new table outputs before saving.
That is not a direct reread of the packed table's N3:N9 cells. If those cells
appear old in the packed copy, identify that document's path and report it
so table persistence can be investigated separately.

### Suppression-time "Save Modified Documents" dialogs

UI mode temporarily sets the template's **loaded, saved references** to
read-only in SolidWorks memory and enables **System Options > External
References > Don't prompt to save read-only referenced documents (discard
changes)** via `swExtRefNoPromptOrSave`. This uses supported API behavior
instead of sending keystrokes or clicking arbitrary dialogs.

The guard discovers recursive dependencies, including suppressed references,
and affects only matching documents already loaded when it starts. The main
template and unrelated open documents are not made read-only. If an affected
reference already has unsaved changes, the workflow stops **before changing
any reference state**: resolve those changes manually and rerun. This also
protects shared references with existing unsaved work in another assembly.

The guard remains active through template updates, Pack-and-Go and closing
the original without saving. Before opening the packed assembly/exporting, it restores the
original discard setting and each still-loaded reference's original
read-only state. Error cleanup attempts every restoration even if one fails,
and reports failures explicitly. The guard does not change filesystem
permissions or file contents and never saves template references.

The design-table adapter refreshes protection before each table commit and
rebuild. It also refreshes protection before closing the original. Previously checked paths that have been
unloaded/reloaded are made read-only again without replacing their original
state snapshot. Newly loaded references are added only if clean; untracked
dirty references stop the workflow rather than being silently discarded.
Virtual components or references loaded within a blocking API call may
still prompt. This is not a universal "Don't Save" handler. Automatic discard
can unload modified suppressed references, so verify the packed model's
geometry as well as its global values. Test this mode on a disposable
template copy before production use.

The discard preference is application-wide while the guard is active:
**do not edit/close other documents or run other macros during UI execution**.
VBA Reset, End, or a SolidWorks crash bypasses cleanup. If that happens, check
the External References discard option and reload affected template documents
without saving to clear temporary in-memory read-only flags.

### Choices and global mapping

The form follows the part-numbering PDFs provided with this task rather than
the initial example labels. Unsupported fifth walls, Both, Open/Closed/U-Shape,
Aluminum, and additional gauges are not invented.

| Form choice | Equation / existing module global | Code |
|---|---|---|
| Wall count 1-4 | `WALL_COUNT` / `wallCount` | 1-4 |
| Right / Left, only when count is 2 | `SIDE` / `side` | 1 / 2 |
| N/A, for counts 1, 3, 4 | `SIDE` / `side` | 0 |
| No return / Single 1 inch / Single 2 inches / Double | `RETURN_TYPE` / `returnType` | 0 / 1 / 2 / 3 |
| Mild Steel / Borcon Weathering Steel | `OVERALL_MATERIAL` / `overallMaterial` | 1 / 2 |
| 3/16 inch / 1/4 inch | Table `N8`; physical `OVERALL_THICKNESS`; naming `thicknessLabel` | 0.1875 / 0.25 inches; naming 1 / 2 |

### Embedded design table inputs

The adapter is based on the supplied `PLANTER_ASSEMBLY.xlsx` workbook.
SolidWorks edits its **embedded** table; the repository workbook is only a
reference and is never opened or saved by the macro.

| Cell | Value written by UI |
|---|---|
| N3 | `ONE WALL`, `TWO WALLS (CENTER AND LEFT SIDE)`, `TWO WALLS (CENTER AND RIGHT SIDE)`, `THREE WALLS`, or `FOUR WALLS` |
| N4 | `Default`, `single_return1`, `single_return2`, or `double_return` |
| N5 | Overall length, numeric inches |
| N6 | Overall width, numeric inches |
| N7 | Overall height, numeric inches |
| N8 | Text `3/16` or `1/4` (never an Excel date) |
| N9 | `MILD STEEL` or `BORCON WEATHERING STEEL` |

Only N3:N9 are inputs. **N10 and N11 are calculated by Excel**, not written
by the form; their results drive the center/side planter globals and naming.
The calculated center/side length fields are not shown in the input form.
**N12 is not used for macro naming and is not edited.**
Existing helper/output formulas and thickness/material rules are preserved.

Enter the three overall dimensions as decimal **inches**, such as `100.25`,
not fractions, unit suffixes, scientific notation or comma decimals. The
workbook's actual data validations allow **6 <= height <= 48**, with length
and width each at least 6 and strictly greater than height. Those formulas,
not the more restrictive text notes in column O, define UI validation.
The template's document length units must be inches.
Naming continues to use the existing `FormatPNDimension` rounding.

`quoteOnly`, `exportPDFAssemblies`, `exportPDFComponents`, `exportDXF`,
`exportSTEP`, and `outputFolder` are public in the new module. Quote-only
clears every export and returns standalone quantities; the original manual
macro's quote/PDF behavior is unchanged. Unchecking quote-only re-enables
the individual choices without reselecting any. Without
quote-only, all exports can be off for model-only output.

### Workflow and safeguards

- Start from the saved `planter_assembly.SLDASM` template and select the
  configuration named in the table's A3 cell (`PLANTER ASSEMBLY` in the
  reference workbook) before opening the UI. Close the Excel/design-table
  editor first. Cancel, Escape, and closing the form do not run automation
  or edit the design table.
- For model runs the wrapper reuses Step 1, Steps 3/6/7 and the existing
  naming/export helpers, with instrumented UI collection instead of Step 4,
  but does not call `RunPackAndGo`, the interactive destination/save Steps
  5/8, or `PromptAndRunExports`. All option selections happen in the form.
- The selected parent must already exist. Output goes to
  `<parent>\PACK_TEST\<part-number>\`, unless the parent already contains a
  complete `PACK_TEST` path segment. The unchanged exporters require that
  marker. Existing part-number folders are rejected, never cleared or
  overwritten; select another parent for another run of the same part.
- The adapter opens the embedded design table using `EditTable2(True)`,
  validates the headers, input labels, formula cells and configuration, and
  preserves the output formulas while writing N3:N9. Excel recalculates the existing formulas;
  SolidWorks commits the table and performs **two consecutive full rebuilds**
  before preview verification or packing, allowing dependent geometry to
  settle on the second pass. Each pass refreshes reference protection, checks
  the configuration/API result, and is timed separately. A failure in either
  pass stops the workflow. Direct writes
  to table-owned equation globals have been removed.
  In preview and full UI runs, only that workbook's Excel windows are hidden
  while inputs are edited and calculated; other workbooks and Excel application
  visibility are left unchanged. Original window visibility is restored and
  window references released before SolidWorks commits/closes the editor,
  including failure cleanup. Excel may briefly appear when opening/closing
  the editor, and later SolidWorks operations may independently show it.
  This is not a fully silent Excel mode. Install this behavior by replacing
  the `UIDesignTableSession` class, then compiling and saving the macro.
  Rebuild checks the active configuration first; it requests a configuration
  switch only when necessary and verifies the resulting active name before
  rebuilding. It does not try to reactivate an already-active configuration.
  SolidWorks API success checks compare explicitly to `False`: the API can
  return True as numeric 1, for which VBA's bitwise `Not` is still nonzero.
  This avoids falsely reporting successful rebuilds, configuration changes,
  table commits or saves as failures.
- Microsoft Excel must be installed. No Excel reference is required in VBA:
  the worksheet is late-bound. The adapter releases the worksheet before
  asking SolidWorks to close the table, and never calls Excel `Quit` or
  closes unrelated workbooks. Externally linked tables and incompatible
  table layouts are explicitly rejected rather than silently edited.
- Calculated globals are checked against the actual rebuilt model, including
  physical `OVERALL_THICKNESS`. Naming's `thicknessLabel` uses the code
  matching the table gauge, not a stale `THICKNESS_LABEL` model value.
  Material assignment and component suppression still depend on the
  template's existing rules; no new material database assignment is invented.
- After every Pack-and-Go save status succeeds and the packed main file is
  confirmed, the original template is **closed without saving** with
  `ISldWorks.CloseDoc`. There is no original-table restoration or second
  template rebuild. The original disk files are not saved by this workflow.
  Stage 7 now closes the original and releases temporary reference protection.
- Pre-existing unsaved main-template changes block the run, as do unsaved
  affected references. Reopen a clean template before starting. On failure,
  the table editor is closed if possible, temporary protection/settings are
  restored, and any still-open original is left unsaved for inspection:
  **close it without saving** to discard changes. Do not save its run-time
  table edits back into the template.
- `CloseDoc` can also unload hidden documents; do not keep unrelated unsaved
  work open during automation. References held by other open documents may
  remain loaded; the workflow reports if the original itself stays loaded.
  Separately opened template references can remain dirty and unsaved; review
  or close them without saving before the next run.
- Native document collection/renaming and every Pack-and-Go save status
  must succeed. The manual fallback is rejected in UI mode because it
  does not implement the requested naming. The packed assembly is opened
  in the same configuration, rebuilt, checked against the calculated
  table outputs, and saved before the selected export macros are invoked.
- Selected macro files are checked before execution, and the packed assembly
  is reactivated before each export. The unchanged exporters still display
  status, cancellation, completion, and error dialogs. Their entry points
  return no export-result status, so a successful invocation is **not**
  presented as proof that every export succeeded; review their dialogs/logs.
  A completely unattended export run is not possible without changing them.
- Errors stop the workflow visibly. Partial output is kept for inspection;
  nothing is automatically deleted.

### Validation

Run `python -m unittest discover -s tests -v` from the repository root.
The dependency-free tests check source wiring and execute the validation,
mapping, table-label, dimension-validation and folder-normalization helper bodies under Windows Script Host
using a type-stripping VBScript adapter. They cover all 80 supported
selection combinations, exact PDF-numbering examples using the existing
naming helper, decimal boundary/invalid inputs, quote-only wiring, and
export routing. This is **not** a VBA compile test or a SolidWorks
integration test.

With Excel installed and the reference workbook present, run the optional
real-Excel formula checks in PowerShell:

```powershell
$env:PLANTER_TEST_EXCEL = '1'
python -m unittest discover -s tests -v
```

These exercise 240 combinations of selections and dimensions in an isolated
Excel instance with the workbook opened read-only. All original cell
contents/formulas are compared after restoration, and the workbook's file
hash must remain unchanged. They do not replace live SolidWorks acceptance
checks.

After **Debug > Compile** in SolidWorks, use a disposable copy of the template
for these acceptance checks:

1. Cancel/Escape/close: no output and no design table changes.
2. Invalid dimensions or missing output folder: stay on the form with an
   actionable error. Wall count 2 requires Left/Right; other counts use N/A.
3. Model only: verify geometry, component suppression and assigned material
   against the inputs, not merely filenames. For four walls, no returns,
   100-inch length/width, 10-inch height, 3/16-inch Mild Steel, the main name
   must be `40-0-C10000-S10000-1000-11.SLDASM`.
4. Verify the original closes without saving, reopen it and confirm its
   disk-backed table inputs are unchanged;
   confirm packed table inputs and calculated globals retain the submitted
   values after reopen. Verify N10:N11 formulas are unchanged. For the
   workbook's original 500 x 500 x 48-inch, four-wall, 1/4-inch setup,
   center length is 113 and side length is 100.4.
5. Quote-only: with no document open, verify the report matches the reference
   example above and no model/Excel/folder/export activity occurs. Also test
   while an unrelated or dirty document is active: it must remain untouched.
   Test the 120.5-inch discrepancy warning. Then test the four export choices
   independently and together in normal model mode, with `.swp` files deployed.
6. Repeat a part number in the same parent: fail without overwriting.
   Check permissions, unavailable macros and incompatible table layouts produce
   explicit errors, not success messages.

The local checks cannot verify SolidWorks table commits, live geometry,
component suppression or actual exports without the template and deployed
macro projects. If the original template has rebuild errors, the adapter
stops before editing its inputs; review those feature/component errors first.

---

Target location: `Z:\Engineering\1_PDM VAULT\SOLIDWORKS DATABASES\MACROS\PLANTER_WALLS\FINAL\`

This folder mirrors that target. Each `.bas` file now has a header comment
at the top ("DEPLOY TARGET") stating exactly which live `.swp` file and
module it goes into, and what action to take - so the instructions travel
with the file even without this README.

**Why `.bas` and not `.swp`:** `.swp` files are compiled VBA project
binaries that only SolidWorks itself can create/write correctly. I can't
safely synthesize that binary format, so these are plain-text source files
for you to paste into the VBA editor (Tools > Macro > Edit / New).

## 1. packngo.bas -> packngo.swp (module: packngo1)
**CHANGED.** In `Step8_ConfirmAndSave`:
- Now captures the newly opened packed assembly and calls `ForceRebuild3`
  on it (previously it opened it but never rebuilt).
- After rebuilding, calls the new `PromptAndRunExports` sub, which asks
  "Is this just for a quote?" and, if not, asks separately about DXF/STEP/PDF
  files, then runs the matching macro(s) below via `swApp.RunMacro2`.
- Added `PromptAndRunExports`, `DeriveModuleName`, `RunExternalMacro` helpers
  and `MACRO_FOLDER` constant near the top (shared state section).

## 2. pdfassemblies.bas -> pdfassemblies.swp (module: pdfassemblies1)
**CHANGED - IMPORTANT.** The live version currently uses old fixed
`SetPosition` coordinates for the Top/Front/Isometric views in
`ProcessMainAssemblyViews`, which is exactly the bug that caused views to
visually shift at different scales. This version replaces that with:
- Outline-based centering (measures the view's actual visual bounding box).
- An anchor-to-outline offset correction (the root cause fix - `View.Position`
  is an internal anchor point, not the bounding box corner, and the gap
  between them scales with view scale).
- New measured center coordinates: Top (5.31, 6.27), Front (5.31, 2.78),
  Isometric (11.29, 8.51) in inches - captured via GET_VIEW_POSITIONS.bas.

## 3. pdfcomponents.bas -> pdfcomponents.swp (module: pdfcomponents1)
**UNCHANGED** from the live version (already had the correct offset-based
positioning). Included here only for completeness/reference.

## 4. dxfcomponents.bas -> NEW FILE
Does not exist live yet. In SolidWorks: Tools > Macro > New, save as
`dxfcomponents.swp` in the FINAL folder above, then paste this file's
content into its module. Exports component (part) drawings to DXF -
same view positioning/scaling logic as pdfcomponents, just saved as .DXF
into a new `DXF\` folder instead of `PDF\`.

## 5. stepcomponents.bas -> NEW FILE
Does not exist live yet. Same process: Tools > Macro > New, save as
`stepcomponents.swp` in the FINAL folder. Exports active, non-suppressed
PART components (not assemblies) directly from the open assembly to STEP,
using `SaveAs` with `swSaveAsOptions_SaveCopy` so it doesn't disturb the
assembly's live references to those parts.

## After importing all 5
Double check in the VBA editor's Project Explorer that each file's module
is still actually named `<basefilename>1` (e.g. `dxfcomponents1`) after you
save - that's the naming convention `RunExternalMacro`/`DeriveModuleName`
in packngo.bas relies on to call these macros automatically. If SolidWorks
names a new macro's module differently on your version, update
`DeriveModuleName` in packngo.bas to match.
