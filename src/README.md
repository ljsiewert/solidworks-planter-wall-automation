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
   restores the original inputs; no Excel VBA-project access is needed.
   Also import [UIReferenceSaveGuard.cls](UIReferenceSaveGuard.cls), which
   protects loaded template references against suppression-time save prompts.
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
The window closes on completion or failure; failures still restore the
template and show the existing error dialogs. Timing-storage errors are
reported separately without claiming the model workflow failed.

**Wait for the current run to finish before updating/recompiling the macro.**
These additions cannot attach a progress window to an already-running macro.
While a modeless progress window is displayed, do not change the active
document/configuration or start another macro; repaint/event processing is
for visibility, not permission to edit the model mid-workflow.

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

The guard remains active through template updates, Pack-and-Go and template
restoration. Before opening the packed assembly/exporting, it restores the
original discard setting and each still-loaded reference's original
read-only state. Error cleanup attempts every restoration even if one fails,
and reports failures explicitly. The guard does not change filesystem
permissions or file contents and never saves template references.

References loaded for the first time during the run are not pre-protected;
dialogs for those documents, virtual components or other save operations may
still appear. This is not a universal "Don't Save" handler. Automatic discard
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
  snapshots N3:N9 before writing. Excel recalculates the existing formulas;
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
- The original N3:N9 contents and N8 formatting are restored after
  Pack-and-Go, including on failure. The restored table calculations and
  rebuilt model globals are verified. The original document stays open
  and is **not saved or closed**, preserving pre-existing unsaved work.
  Table changes/restoration
  can leave it marked dirty; inspect it before choosing to save. Failed
  restoration is reported explicitly.
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
4. Verify original table inputs are restored and no template file is saved;
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
