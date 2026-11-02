"""Preview routing contracts and actual checkbox event behavior on a fake form."""
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
MODULE = (ROOT / "src" / "userform_automation.bas").read_text(encoding="ascii")
FORM = (ROOT / "src" / "UserForm_AutomationUI.frm").read_text(encoding="ascii")


class PreviewTests(unittest.TestCase):
    def test_preview_isolated_from_output_operations(self):
        preview = MODULE.split("Private Sub UIRunPreview()", 1)[1].split("End Sub", 1)[0]
        for forbidden in ("SavePackAndGo", "UICheckDestination", "UICreateDestination",
                          "UIRunExport", "CloseDoc", "Save3", "OpenDoc6"):
            self.assertNotIn(forbidden, preview)
        self.assertNotIn("UIRequireCleanTemplate", preview)
        self.assertNotIn("GetSaveFlag", preview)
        self.assertIn("tableSession.Apply template, designTableInputs, expected", preview)
        self.assertIn("UIReadCalculatedGlobals template, expected", preview)
        self.assertIn("saveGuard.Restore", preview)
        self.assertIn("UICloseTableEditor(tableSession)", preview)
        self.assertIn("UIRestoreSaveGuard(saveGuard)", preview)
        self.assertIn('runLogger.Finish "PREVIEW"', preview)
        self.assertIn("If previewOnly Then\n                UIRunPreview\n            Else", MODULE)
        self.assertIn("If Not previewBox.Value And Not quoteBox.Value Then Call UIOutputRoot(selectedFolder)", FORM)
        self.assertEqual(preview.count('plan.Add "'), 2)
        self.assertEqual(preview.count("progress.NextStage"), 2)

    @unittest.skipUnless(shutil.which("cscript"), "Windows Script Host required")
    def test_checkbox_transitions_and_defensive_export_reset(self):
        methods = []
        for source, name in ((FORM, "previewBox_Click"), (FORM, "quoteBox_Click"),
                             (FORM, "ApplyModeLocks"), (FORM, "exportsOnlyBox_Click"),
                             (MODULE, "UIDisableExports")):
            body = re.search(rf"^(?:Private|Public) Sub {name}\(\).*?^End Sub$",
                             source, re.MULTILINE | re.DOTALL)[0]
            body = body.replace(" As Boolean", "")
            methods.append(body.replace("Private Sub", "Sub").replace("Public Sub", "Sub"))
        script = r'''
Option Explicit
Dim previewBox, quoteBox, folderBox, browseButton, pdfAssembliesBox, pdfComponentsBox, dxfBox, stepBox, priceBox, gapBox, marginBox
Dim quoteOnly, exportPDFAssemblies, exportPDFComponents, exportDXF, exportSTEP
Dim fieldDesigns, quantityBox, designNameBox, addPlanterButton, removePlanterButton, updatePlanterButton, newPlanterButton, planterList, runButton
Dim exportsOnlyBox, rerunPath, customBox, wallsBox, sideBox, returnBox, lengthBox, widthBox, heightBox
Dim materialBox, thicknessBox, mountBox, overrideBox, baseWidthBox, item
Class Control
    Public Value, Enabled, Caption, ListIndex
End Class
Class Designs
    Public Count
End Class
Sub Assert(condition, message)
    If Not condition Then
        WScript.Echo "FAIL: " & message
        WScript.Quit 1
    End If
End Sub
Set previewBox = New Control
Set quoteBox = New Control
Set folderBox = New Control
Set browseButton = New Control
Set pdfAssembliesBox = New Control
Set pdfComponentsBox = New Control
Set dxfBox = New Control
Set stepBox = New Control
Set priceBox = New Control
Set gapBox = New Control
Set marginBox = New Control
Set fieldDesigns = New Designs
fieldDesigns.Count = 0
Set quantityBox = New Control
Set designNameBox = New Control
Set addPlanterButton = New Control
Set removePlanterButton = New Control
Set updatePlanterButton = New Control
Set newPlanterButton = New Control
Set planterList = New Control
planterList.ListIndex = -1
Set runButton = New Control
Set exportsOnlyBox = New Control
exportsOnlyBox.Value = False
rerunPath = ""
Set customBox = New Control
Set wallsBox = New Control
Set sideBox = New Control
Set returnBox = New Control
Set lengthBox = New Control
Set widthBox = New Control
Set heightBox = New Control
Set materialBox = New Control
Set thicknessBox = New Control
Set mountBox = New Control
Set overrideBox = New Control
Set baseWidthBox = New Control
wallsBox.Value = "2"
overrideBox.Value = False
'''
        script += "\n".join(methods) + r'''
quoteBox.Value = True
pdfAssembliesBox.Value = True
pdfComponentsBox.Value = True
dxfBox.Value = True
stepBox.Value = True
previewBox.Value = True
previewBox_Click
Assert Not quoteBox.Value And Not pdfAssembliesBox.Value And Not pdfComponentsBox.Value, "Preview clears PDFs/quote"
Assert Not dxfBox.Value And Not stepBox.Value, "Preview clears DXF/STEP"
Assert Not quoteBox.Enabled And Not pdfAssembliesBox.Enabled And Not stepBox.Enabled, "Preview disables export choices"
Assert Not folderBox.Enabled And Not browseButton.Enabled, "Preview needs no destination"
previewBox.Value = False
previewBox_Click
Assert quoteBox.Enabled And pdfAssembliesBox.Enabled And folderBox.Enabled, "Exiting preview reenables controls"
Assert Not pdfAssembliesBox.Value And Not stepBox.Value, "Exiting preview does not restore hidden selections"
quoteBox.Value = True
quoteBox_Click
Assert Not pdfAssembliesBox.Value And Not pdfAssembliesBox.Enabled, "Quote disables all PDFs"
Assert Not folderBox.Enabled And Not browseButton.Enabled And Not previewBox.Enabled, "Quote needs no folder or preview"
Assert Not runButton.Enabled, "Empty field disables Run"
fieldDesigns.Count = 1
quoteBox_Click
Assert runButton.Enabled And runButton.Caption = "Run field", "Added design enables field run"
fieldDesigns.Count = 0
quoteBox.Value = False
quoteBox_Click
Assert folderBox.Enabled And previewBox.Enabled And pdfAssembliesBox.Enabled, "Leaving quote reenables controls"
Assert Not pdfAssembliesBox.Value, "Leaving quote does not reselect PDFs"
Assert Not exportsOnlyBox.Enabled, "Exports only needs a rerun source"
rerunPath = "Z:\out\20-0-C10000-S6000-1200-11-C1\20-0-C10000-S6000-1200-11-C1.SLDASM"
quoteBox_Click
Assert exportsOnlyBox.Enabled And Not previewBox.Enabled And folderBox.Enabled And browseButton.Enabled, "Rerun locks preview but allows destination"
Assert Not customBox.Enabled And wallsBox.Enabled And sideBox.Enabled And Not baseWidthBox.Enabled, "Rerun keeps model inputs editable"
exportsOnlyBox.Value = True
exportsOnlyBox_Click
For Each item In Array(quoteBox, wallsBox, sideBox, returnBox, lengthBox, widthBox, heightBox, materialBox, thicknessBox, mountBox, overrideBox, baseWidthBox)
    Assert Not item.Enabled, "Exports only locks model inputs"
Next
Assert pdfAssembliesBox.Enabled And dxfBox.Enabled And stepBox.Enabled, "Exports only keeps export choices"
Assert Not folderBox.Enabled And Not browseButton.Enabled, "Exports only locks destination"
exportsOnlyBox.Value = False
exportsOnlyBox_Click
Assert quoteBox.Enabled And lengthBox.Enabled And Not previewBox.Enabled, "Leaving exports only unlocks inputs"
Assert folderBox.Enabled And browseButton.Enabled, "Leaving exports only unlocks destination"
quoteBox.Value = True
quoteBox_Click
Assert Not exportsOnlyBox.Enabled, "Quote disables exports only"
quoteBox.Value = False
quoteBox_Click
rerunPath = ""
quoteOnly = True
exportPDFAssemblies = True
exportPDFComponents = True
exportDXF = True
exportSTEP = True
UIDisableExports
Assert Not quoteOnly And Not exportPDFAssemblies And Not exportPDFComponents And Not exportDXF And Not exportSTEP, _
    "Defensive reset forces all exports off"
WScript.Echo "PASS: preview/quote transitions and export reset"
'''
        with tempfile.TemporaryDirectory(prefix="planter-preview-") as folder:
            runner = Path(folder) / "preview.vbs"
            runner.write_text(script, encoding="ascii")
            result = subprocess.run(["cscript", "//Nologo", str(runner)], capture_output=True,
                                    text=True, timeout=30)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertIn("PASS:", result.stdout, result.stdout + result.stderr)
