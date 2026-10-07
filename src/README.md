# Deploy These Macro Updates to SolidWorks

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
