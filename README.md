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
  planter lengths are calculated by the existing design table.
- Mild Steel or Borcon Weathering Steel, in 3/16-inch or 1/4-inch thickness.
- Output folder selection, quote-only mode, and independent assembly PDF,
  component PDF, DXF and STEP selections.
- A progress window with current stage, completed milestones, elapsed time
  and rough remaining-time estimates from previous successful runs.
- Local per-run timing logs for investigating Excel, rebuild, Pack-and-Go
  and exporter bottlenecks; logs are kept out of Git.
- A scoped read-only reference guard intended to discard suppression-time
  template changes without saving the originals.

## Install and run

Requires Windows, SolidWorks with VBA support, and Microsoft Excel for the
embedded design table.

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