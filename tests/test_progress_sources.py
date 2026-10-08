"""Portable progress logic tests; SolidWorks/MSForms rendering needs live testing."""

from itertools import product
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
MODULE = (ROOT / "src" / "userform_automation.bas").read_text(encoding="ascii")
SESSION = (ROOT / "src" / "UIProgressSession.cls").read_text(encoding="ascii")
FORM = (ROOT / "src" / "UserForm_AutomationProgress.frm").read_text(encoding="ascii")


class ProgressSourceTests(unittest.TestCase):
    def test_all_export_plans_match_workflow(self):
        plan = MODULE.split("Private Function UIProgressPlan()", 1)[1].split("End Function", 1)[0]
        workflow = MODULE.split("Private Sub UIRunAutomation()", 1)[1].split("Failed:", 1)[0]
        flags = ("exportDXF", "exportSTEP", "exportPDFAssemblies", "exportPDFComponents")
        base_stages = re.findall(r'^\s+plan.Add "([^"]+)"', plan, re.MULTILINE)
        self.assertEqual(len(base_stages), 9)
        self.assertEqual(workflow.count("progress.NextStage"), 13)
        for choices in product((False, True), repeat=4):
            stage_count = len(base_stages)
            advances = workflow.count("progress.NextStage")
            for flag, selected in zip(flags, choices):
                self.assertRegex(plan, rf'If {flag} Then plan.Add "[^"]+"')
                branch = re.search(rf"If {flag} Then(.*?)End If", workflow, re.DOTALL).group(1)
                self.assertEqual(branch.count("progress.NextStage"), 1)
                self.assertLess(branch.index("progress.NextStage"), branch.index("UIRunExport"))
                stage_count += int(selected)
                advances -= int(not selected)
            self.assertEqual(stage_count, advances)
        for operation in ("tableSession.Apply", "SavePackAndGo", "CloseDoc templatePath",
                          "OpenDoc6", "packed.ForceRebuild3"):
            self.assertIn(operation, workflow)
        self.assertLess(workflow.index("progress.Complete"), workflow.index('MsgBox "Pack-and-Go'))

    def test_display_and_failure_contracts(self):
        self.assertIn("form.Show vbModeless", SESSION)
        self.assertIn("DateDiff", SESSION)
        self.assertNotIn("Timer", SESSION)
        self.assertIn("GetSetting", SESSION)
        self.assertIn("SaveSetting", SESSION)
        self.assertIn("remaining = -1", SESSION)
        self.assertIn("On Error GoTo Failed", SESSION)
        self.assertIn("Me.Repaint", FORM)
        self.assertIn("Elapsed at last update:", FORM)
        self.assertIn("Cancel = True", FORM)
        self.assertIn("If Not progress Is Nothing Then progress.CloseWindow", MODULE)
        self.assertIn("If current < 1 Then Exit Sub", SESSION)


@unittest.skipUnless(shutil.which("cscript"), "Windows Script Host required")
class ProgressBehaviorTests(unittest.TestCase):
    def test_stage_counts_eta_and_completion(self):
        procedures = []
        for name in ("BeginRun", "NextStage", "Complete", "CloseWindow"):
            body = re.search(rf"^Public Sub {name}\b.*?^End Sub$", SESSION,
                             re.MULTILINE | re.DOTALL).group(0)
            body = re.sub(r" As (?:Collection|Long|Double|String)\b", "", body)
            body = re.sub(r"^(\s*Next) \w+$", r"\1", body, flags=re.MULTILINE)
            body = body.replace("Unload form", "form.MarkClosed")
            body = body.replace("ReDim durations(1 To stages.Count)", "ReDim durations(stages.Count)")
            procedures.append(body)
        script = r'''
Option Explicit
Const SETTINGS_APP = "Test"
Const SETTINGS_SECTION = "Test"
Const vbModeless = 0
Dim form, stages, durations(), current, started, stageStarted
Dim history, plan, display, saveCalls, number
Set history = CreateObject("Scripting.Dictionary")
saveCalls = 0
Class FakePlan
    Public Property Get Count()
        Count = 2
    End Property
    Public Default Property Get Item(index)
        Item = "stage" & index
    End Property
End Class
Class UserForm_AutomationProgress
    Public completed, total, remaining, elapsed, closed
    Public Sub Show(mode)
        closed = False
    End Sub
    Public Sub UpdateProgress(stage, c, t, e, r)
        completed = c
        total = t
        elapsed = e
        remaining = r
    End Sub
    Public Sub MarkClosed()
        closed = True
    End Sub
End Class
Function GetSetting(app, section, key, fallback)
    If history.Exists(key) Then
        GetSetting = history(key)
    Else
        GetSetting = fallback
    End If
End Function
Sub SaveTimings()
    saveCalls = saveCalls + 1
End Sub
Sub DoEvents()
End Sub
Sub UITraceBegin(name)
End Sub
Sub UITraceEnd(name)
End Sub
Sub Assert(condition, message)
    If Not condition Then
        WScript.Echo "FAIL: " & message
        WScript.Quit 1
    End If
End Sub
'''
        script += "\n" + "\n\n".join(procedures) + r'''
Set plan = New FakePlan
BeginRun plan
Set display = form
NextStage
Assert display.completed = 0 And display.total = 2, "Begin at 0/2"
Assert display.remaining = -1, "First run has no invented ETA"
history.Add "stage1", "40"
history.Add "stage2", "60"
CloseWindow
Assert display.closed, "Window closes on failure"
Assert saveCalls = 0, "Failed run does not save timings"
BeginRun plan
Set display = form
NextStage
Assert display.remaining = 100, "Sum all remaining recorded stages"
NextStage
Assert display.completed = 1 And display.remaining = 60, "Exclude completed stage from ETA"
Complete
Assert display.completed = 2 And display.remaining = 0, "Complete only after last stage"
Assert display.closed And saveCalls = 1, "Successful workflow closes and saves timings"
BeginRun plan
NextStage
On Error Resume Next
Complete
number = Err.Number
Err.Clear
On Error GoTo 0
Assert number <> 0, "Early completion is an error"
CloseWindow
history.Remove "stage2"
BeginRun plan
NextStage
Assert form.remaining = -1, "Missing a future stage timing means unknown ETA"
CloseWindow
WScript.Echo "PASS: stage counts, first-run ETA, historical ETA, completion and failure cleanup"
'''
        with tempfile.TemporaryDirectory(prefix="planter-progress-test-") as directory:
            runner = Path(directory) / "progress.vbs"
            runner.write_text(script, encoding="ascii")
            result = subprocess.run(["cscript", "//Nologo", str(runner)],
                                    capture_output=True, text=True, timeout=30)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertIn("PASS:", result.stdout, result.stdout + result.stderr)
