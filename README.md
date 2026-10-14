# SolidWorks Planter Wall Automation

VBA automation for configuring planter-wall assemblies, creating renamed
Pack-and-Go copies, and generating manufacturing documentation.

## Two entry points

- **Manual:** `packngo1.RunPackAndGo` preserves the existing prompt-driven
  workflow.
- **UI:** `userform_automation.RunWithUI` collects the design and export
  selections before starting automation. It updates the embedded design
  table instead of writing to its controlled equation globals.

The five original macro sources remain unchanged by the UI integration.

## UI features

- One to four walls, with Left/Right orientation for two-wall layouts.
- No return, single 1-inch return, single 2-inch return, or double return.
- Overall length, width and height in decimal inches; center and side
  planter lengths are calculated by the design table for model runs or by
  equivalent standalone formulas for quote-only.
- Mild Steel or Borcon Weathering Steel, in 3/16-inch or 1/4-inch thickness.
- Output folder selection, quote-only mode, and independent assembly PDF,
  component PDF, DXF and STEP selections.
- Quote-only returns input-only component dimensions, quantities and profile
  areas using the supplied component formulas, without opening a model,
  Excel, rebuilding, Pack-and-Go, exports or an output folder. Profile weight
  and material cost use a shared estimating steel density and an editable
  $/lb input defaulting to 0.90. All exports start unchecked.
- Quote pricing adds labor (cutting per sheet, bending per bent part when a
  return is selected, welding per part), optional paint (Painted toggle next
  to Material; $/sq ft on both sides), total cost, gross margin and sell
  price. Material is priced on full purchased sheets so scrap is included.
  Rates, material $/lb, part gap and edge margin live under **Properties...**
  on the input form; margin can also be adjusted live on the results page.
  Default gross margin is 50%; welding is $10 per part. Operations' part-gap
  baseline is 0.25 in, with zero available for touching-edge estimates.
- Quote results open with a cost/weight/design summary and a compact component
  table. Technical Details retains the full calculation report and assumptions.
- A 120 x 60-inch sheet estimator nests rectangles and shape-aware opposed
  gusset pairs, with 90-degree rotation and editable gap/margin assumptions.
  It compares 24 split-based layouts and eight MaxRects layouts automatically,
  then tries single/mixed gussets and least-used-sheet elimination, keeping
  the lowest sheet count without regressing the earlier estimator.
  It tries 24 deterministic guillotine layouts and reports the best estimate
  for sheets, utilization and unused/scrap percentage. Operations verification
  is required, and no laser toolpaths are generated.
- **Preview Sheet Layout**, next to Technical Details, opens a PDF in the
  default viewer with one stock sheet per page, component outlines and counts,
  and utilization percentages. This is not a production cutting plan.
- Results include a per-sheet **Unused / potential scrap** table and total,
  distinguishing unused stock from material that must be discarded.
- Preview-only selection updates/rebuilds the unsaved template without
  Pack-and-Go or exports; export choices are cleared and disabled. Close
  without saving and reopen before the next run. Calculated center/side
  lengths are no longer displayed as input fields.
- A progress window with current stage, completed milestones, elapsed time
  and rough remaining-time estimates from previous successful runs.
- Local per-run timing logs for investigating Excel, rebuild, Pack-and-Go
  and exporter bottlenecks; logs are kept out of Git.
- A scoped read-only reference guard intended to discard suppression-time
  template changes without saving the originals.

## Install and run

Requires Windows and SolidWorks with VBA support. Microsoft Excel and a clean
template are required for model/preview runs, not input-only quoting.

1. Back up the existing macro and use a disposable template copy for testing.
2. Edit the existing Pack-and-Go `.swp` project in SolidWorks.
3. Import [userform_automation.bas](src/userform_automation.bas) and the
   [design-table](src/UIDesignTableSession.cls),
   [progress](src/UIProgressSession.cls), and
   [reference-save guard](src/UIReferenceSaveGuard.cls) class modules.
   Also import the [runtime logger](src/UIRunLogger.cls).
4. Create UserForms named `UserForm_AutomationUI` and
   `UserForm_AutomationProgress`. Paste their respective
   [input-form](src/UserForm_AutomationUI.frm) and
   [progress-form](src/UserForm_AutomationProgress.frm) code into the code
   windows. These `.frm` files are code-behind, **not directly importable
   native form packages**; controls are created at runtime.
5. Compile and save the VBA project. Keep the original module named
   `packngo1`; keep the four exporters as separate `.swp` projects.
6. Open `planter_assembly.SLDASM`, activate the configuration listed in the
   embedded table's A3 cell, close the table editor, and run `RunWithUI`.
7. Enter inputs and choose an existing output parent folder. For the first
   run, leave every export unchecked and verify the packed geometry.

See the **[complete deployment guide](src/README.md)** for references,
installation/update steps, cell mappings, export locations and acceptance
checks.

## Design-table contract

The embedded table's seven inputs are N3:N9: wall layout, return type,
overall length, overall width, overall height, thickness and material.
N10:N11 remain calculated planter lengths. N12 is not used for macro
naming. The adapter validates the expected headers and layout before editing.

Height must be 6-48 inches inclusive; length and width must each exceed
height. Document length units must be inches. Externally linked or
incompatible tables are rejected.

Reference workbooks and proprietary part-numbering PDFs are kept local and
are not distributed in this repository. Supply an authorized template with
the documented table layout to run the automation.

## Safeguards and limits

- Existing part-number output folders are rejected, not overwritten.
- Output uses a `PACK_TEST` path segment required by the unchanged exporters.
- The original template is closed without saving after Pack-and-Go succeeds;
  its design table is not restored or rebuilt back to the original design.
  Temporary read-only/settings changes are restored before opening the copy.
- Pre-existing unsaved main-template changes block the run. On failure,
  any still-open template is left unsaved for inspection; close it without
  saving to discard the run's changes.
- Pre-existing unsaved changes in affected references block automatic
  discard. Newly loaded references may still show save prompts.
- Do not edit other documents or run another macro during automation.
  VBA Reset/End or a SolidWorks crash can bypass cleanup; see the guide for
  recovery.
- Progress updates between stages, not during blocking API calls. Percent
  complete measures milestones rather than files or time. The first run
  has no timing estimate.
- Exporters still display dialogs and do not return per-file success.
  Always inspect export results and verify geometry, suppression and material.

## Tests and validation status

See the deployment guide's [runtime logging section](src/README.md#runtime-logs-and-performance-investigation)
for the log folder, configuration and baseline profiling steps.

From the repository root:

```powershell
python -m unittest discover -s tests -v
```

Dependency-free local tests cover mappings, naming, validation, configuration
and Boolean-return regressions, progress estimates and save-guard cleanup
against portable/fake hosts. They do not compile VBA or validate live
SolidWorks geometry.

An optional test exercises the local reference workbook through installed
Excel, read-only, across 240 scenarios:

```powershell
$env:PLANTER_TEST_EXCEL = '1'
python -m unittest discover -s tests -v
```

That test requires an authorized local `PLANTER_ASSEMBLY.xlsx` and skips when
the workbook or opt-in flag is absent. Live SolidWorks testing is still
required, especially for the progress window and suppression-time save guard.