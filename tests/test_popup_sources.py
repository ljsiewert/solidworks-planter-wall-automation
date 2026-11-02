"""Routine notifications stay nonmodal; decisions and failures remain visible."""
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import unittest

from test_userform_sources import helper_body

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
EXPORTERS = {
    "dxfcomponents": "RunDXFComponentDrawings",
    "stepcomponents": "RunSTEPComponentExport",
    "pdfassemblies": "RunPDFAssemblyDrawings",
    "pdfcomponents": "RunPDFComponentDrawings",
}


class PopupSourceTests(unittest.TestCase):
    def test_routine_messages_are_nonmodal_but_decisions_and_failures_remain(self):
        for module, entry in EXPORTERS.items():
            with self.subTest(module=module):
                source = (SRC / f"{module}.bas").read_text(encoding="utf-8")
                body = helper_body(entry, source)
                self.assertNotIn("vbInformation", body)
                self.assertIn("SetStatusBarText", body)
                self.assertIn('If failureCount > 0 Then\n', body)
                self.assertIn("vbExclamation", body)
                self.assertIn("vbYesNo", source)
                self.assertIn('If CancelRequested() Then', body)
                self.assertIn('MsgBox "No document is open."', body)

    def test_ui_success_paths_are_nonmodal_and_rerun_warnings_are_preserved(self):
        module = (SRC / "userform_automation.bas").read_text(encoding="ascii")
        rerun = (SRC / "ui_rerun.bas").read_text(encoding="ascii")
        progress = (SRC / "UserForm_AutomationProgress.frm").read_text(encoding="ascii")
        self.assertNotIn("vbInformation", module + rerun + progress)
        self.assertIn("Do not save the template; close without saving", module)
        for procedure, source in (
            ("UIRunPreview", module),
            ("UIRunAutomation", module),
            ("UIRunExportsOnly", rerun),
            ("UIRunRerun", rerun),
        ):
            body = helper_body(procedure, source.replace("Private Sub ", "Sub "))
            self.assertIn("SetStatusBarText", body)
        self.assertIn("If Len(openWarnings & cleanupWarning) > 0 Then", rerun)
        self.assertIn('vbExclamation, "Rerun Warnings"', rerun)
        self.assertIn('vbCritical, "Rerun Stopped"', rerun)
        self.assertIn("Cancel = True", progress)
        self.assertIn("stageLabel.Caption", progress)
        self.assertNotIn("MsgBox", progress)
        form = (SRC / "UserForm_AutomationUI.frm").read_text(encoding="ascii")
        self.assertIn('MsgBox Err.Description, vbExclamation, "Review Form Inputs"', form)


@unittest.skipUnless(shutil.which("cscript"), "Windows Script Host required")
class PopupBehaviorTests(unittest.TestCase):
    def test_actual_export_summaries_only_show_popups_for_failures(self):
        for module, entry in EXPORTERS.items():
            with self.subTest(module=module):
                body = helper_body(entry, (SRC / f"{module}.bas").read_text(encoding="utf-8"))
                summary = body[body.index("swApp.Frame.SetStatusBarText IIf(wasCancelled,"):]
                summary = summary.split("    Exit Sub", 1)[0].split("End Sub", 1)[0]
                summary = summary.replace('IIf(wasCancelled,', 'ChooseStatus(wasCancelled,')
                script = r'''
Option Explicit
Const vbExclamation = 48
Dim swApp, frame, wasCancelled, successCount, skippedCount, failureCount
Dim dxfFolder, stepFolder, logFilePath, componentLogPath, popups, cancelled, failed
Class FakeFrame
    Public LastStatus
    Sub SetStatusBarText(text)
        LastStatus = text
    End Sub
End Class
Class FakeApp
    Public Frame
End Class
Sub MsgBox(text, style, title)
    If style <> vbExclamation Then Err.Raise 513, , "Failure must remain a warning"
    popups = popups + 1
End Sub
Function ChooseStatus(condition, yes, no)
    If condition Then
        ChooseStatus = yes
    Else
        ChooseStatus = no
    End If
End Function
Set frame = New FakeFrame
Set swApp = New FakeApp
Set swApp.Frame = frame
successCount = 3
skippedCount = 1
dxfFolder = "output"
stepFolder = "output"
logFilePath = "log"
componentLogPath = "log"
For Each cancelled In Array(False, True)
    wasCancelled = cancelled
    For Each failed In Array(0, 1)
        failureCount = failed
        popups = 0
        ShowSummary
        If popups <> failed Then Err.Raise 513, , "Routine popup or missing failure warning"
        If Len(frame.LastStatus) = 0 Then Err.Raise 513, , "Missing nonmodal result"
        If cancelled And InStr(frame.LastStatus, "cancelled") = 0 Then Err.Raise 513, , "Missing cancellation status"
    Next
Next
'''
                with tempfile.TemporaryDirectory() as folder:
                    path = Path(folder) / "summary.vbs"
                    path.write_text(script + "\nSub ShowSummary\n" + summary + "\nEnd Sub\n", encoding="ascii")
                    result = subprocess.run(
                        ["cscript", "//nologo", str(path)], capture_output=True, text=True, timeout=30
                    )
                    self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
