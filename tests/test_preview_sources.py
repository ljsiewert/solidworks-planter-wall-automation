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
        self.assertIn("UIRequireCleanTemplate template", preview)
        self.assertIn("tableSession.Apply template, designTableInputs, expected", preview)
        self.assertIn("UIReadCalculatedGlobals template, expected", preview)
        self.assertIn("saveGuard.Restore", preview)
        self.assertIn("UICloseTableEditor(tableSession)", preview)
        self.assertIn("UIRestoreSaveGuard(saveGuard)", preview)
        self.assertIn('runLogger.Finish "PREVIEW"', preview)
        self.assertIn("If previewOnly Then\n            UIRunPreview\n        Else", MODULE)
        self.assertIn("If Not previewBox.Value Then Call UIOutputRoot(selectedFolder)", FORM)
        self.assertEqual(preview.count('plan.Add "'), 2)
        self.assertEqual(preview.count("progress.NextStage"), 2)

    @unittest.skipUnless(shutil.which("cscript"), "Windows Script Host required")
    def test_checkbox_transitions_and_defensive_export_reset(self):
        methods = []
        for source, name in ((FORM, "previewBox_Click"), (FORM, "quoteBox_Click"),
                             (MODULE, "UIDisableExports")):
            body = re.search(rf"^(?:Private|Public) Sub {name}\(\).*?^End Sub$",
                             source, re.MULTILINE | re.DOTALL)[0]
            methods.append(body.replace("Private Sub", "Sub").replace("Public Sub", "Sub"))
        script = r'''
Option Explicit
Dim previewBox, quoteBox, folderBox, browseButton, pdfAssembliesBox, pdfComponentsBox, dxfBox, stepBox
Dim quoteOnly, exportPDFAssemblies, exportPDFComponents, exportDXF, exportSTEP
Class Control
    Public Value, Enabled
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
Assert pdfAssembliesBox.Value And Not pdfAssembliesBox.Enabled, "Quote behavior preserved"
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
