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
The two corresponding form fields are locked and display "Calculated by
design table." **N12 is not used for macro naming and is not edited.**
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
forces assembly PDFs on and the other three exports off, matching manual
mode. Unchecking quote-only re-enables the individual choices. Without
quote-only, all exports can be off for model-only output.

### Workflow and safeguards

- Start from the saved `planter_assembly.SLDASM` template and select the
  configuration named in the table's A3 cell (`PLANTER ASSEMBLY` in the
  reference workbook) before opening the UI. Close the Excel/design-table
  editor first. Cancel, Escape, and closing the form do not run automation
  or edit the design table.
- The wrapper reuses Steps 1-4, 6-7 and the existing naming/export helpers,
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
  SolidWorks commits the table and rebuilds **before** packing. Direct writes
  to table-owned equation globals have been removed.
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
5. Quote-only: only assembly PDFs are invoked. Then test the four export
   choices independently and together, with all required `.swp` files deployed.
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
