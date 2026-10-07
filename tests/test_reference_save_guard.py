"""Execute the save guard against a fake COM host; not a SolidWorks integration test."""

from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
GUARD = (ROOT / "src" / "UIReferenceSaveGuard.cls").read_text(encoding="ascii")
MODULE = (ROOT / "src" / "userform_automation.bas").read_text(encoding="ascii")


class ReferenceGuardContracts(unittest.TestCase):
    def test_guard_scope_and_cleanup(self):
        self.assertIn("GetDependencies(True, True, False, False, False)", GUARD)
        self.assertIn("document.GetSaveFlag <> False", GUARD)
        self.assertLess(GUARD.index("document.GetSaveFlag"), GUARD.index("app.SetUserPreferenceToggle"))
        self.assertIn("references.Exists(path)", GUARD)
        self.assertIn("template.GetPathName", GUARD)
        self.assertNotIn("Save3", GUARD)
        self.assertNotIn("SendKeys", GUARD)
        self.assertNotIn("SetAttr", GUARD)
        self.assertIn("issues = issues & RestoreDocument", GUARD)
        self.assertIn("issues = issues & RestorePreference()", GUARD)
        workflow = MODULE.split("Private Sub UIRunAutomation()", 1)[1]
        self.assertLess(workflow.index("saveGuard.BeginGuard"), workflow.index("tableSession.Apply"))
        self.assertLess(workflow.index("tableSession.Restore"), workflow.index("saveGuard.Restore"))
        self.assertLess(workflow.index("saveGuard.Restore"), workflow.index("swApp.OpenDoc6"))
        self.assertIn("UIRestoreSaveGuard(saveGuard)", workflow.split("Failed:", 1)[1])


@unittest.skipUnless(shutil.which("cscript"), "Windows Script Host required")
class ReferenceGuardBehavior(unittest.TestCase):
    def test_protection_preflight_and_restoration(self):
        body = GUARD.split("Option Explicit", 1)[1]
        body = re.sub(r"^Private Const GUARD_ERROR.*$", "", body, flags=re.MULTILINE)
        body = re.sub(r" As (?:SldWorks\.\w+|Collection|Object|String|Long|Boolean|Variant)\b", "", body)
        body = re.sub(r"^(\s*Next) \w+$", r"\1", body, flags=re.MULTILINE)
        body = body.replace("New Collection", "New FakeCollection")
        body = body.replace("swUserPreferenceToggle_e.swExtRefNoPromptOrSave", "15")
        # VBScript has no labeled handlers; preserve each helper's error result contract.
        body = body.replace("On Error GoTo Failed", "On Error Resume Next\n    Err.Clear")
        body = body.replace("    Exit Function\nFailed:", "    If Err.Number = 0 Then Exit Function")
        script = r'''
Option Explicit
Const GUARD_ERROR = -2147219104
Dim app, template, reference, unrelated, guard, number
Class FakeCollection
    Private items, size
    Private Sub Class_Initialize()
        Set items = CreateObject("Scripting.Dictionary")
        size = 0
    End Sub
    Public Sub Add(value)
        size = size + 1
        items.Add CStr(size), value
    End Sub
    Public Property Get Count()
        Count = size
    End Property
    Public Default Property Get Item(index)
        Item = items(CStr(index))
    End Property
End Class
Class FakeExtension
    Public Function GetDependencies(a,b,c,d,e)
        GetDependencies = Array("part", "C:\template\part.SLDPRT")
    End Function
End Class
Class FakeDocument
    Public path, dirty, readOnly, setCalls, failProtect, failRestore, Extension
    Private Sub Class_Initialize()
        dirty = False
        readOnly = False
        setCalls = 0
        failProtect = False
        failRestore = False
        Set Extension = New FakeExtension
    End Sub
    Public Function GetPathName()
        GetPathName = path
    End Function
    Public Function GetSaveFlag()
        GetSaveFlag = dirty
    End Function
    Public Function IsOpenedReadOnly()
        IsOpenedReadOnly = readOnly
    End Function
    Public Function SetReadOnlyState(value)
        setCalls = setCalls + 1
        If (value And failProtect) Or ((value = False) And failRestore) Then
            SetReadOnlyState = False
        Else
            readOnly = value
            SetReadOnlyState = 1
        End If
    End Function
End Class
Class FakeApp
    Public preference, prefCalls
    Private Sub Class_Initialize()
        preference = False
        prefCalls = 0
    End Sub
    Public Function GetDocuments()
        GetDocuments = Array(template, reference, unrelated)
    End Function
    Public Function GetOpenDocumentByName(path)
        If path = reference.path Then
            Set GetOpenDocumentByName = reference
        Else
            Set GetOpenDocumentByName = Nothing
        End If
    End Function
    Public Function GetUserPreferenceToggle(id)
        GetUserPreferenceToggle = preference
    End Function
    Public Sub SetUserPreferenceToggle(id, value)
        prefCalls = prefCalls + 1
        preference = value
    End Sub
End Class
Sub Setup()
    Set app = New FakeApp
    Set template = New FakeDocument
    template.path = "C:\template\PLANTER_ASSEMBLY.SLDASM"
    Set reference = New FakeDocument
    reference.path = "C:\template\part.SLDPRT"
    Set unrelated = New FakeDocument
    unrelated.path = "C:\other.SLDPRT"
    unrelated.dirty = True
    Set guard = New UIReferenceSaveGuard
End Sub
Sub Assert(condition, message)
    If Not condition Then
        WScript.Echo "FAIL: " & message
        WScript.Quit 1
    End If
End Sub
Class UIReferenceSaveGuard
'''
        script += body + r'''
End Class
Setup
guard.BeginGuard app, template
Assert reference.readOnly And app.preference, "Loaded reference protected"
Assert template.setCalls = 0 And unrelated.setCalls = 0, "Main/unrelated documents untouched"
guard.Restore
Assert reference.readOnly = False And app.preference = False, "Original states restored"
guard.Restore
Assert reference.setCalls = 2, "Repeated cleanup does not toggle again"
Setup
reference.readOnly = True
app.preference = True
guard.BeginGuard app, template
guard.Restore
Assert reference.readOnly And app.preference, "Preserve original read-only and discard states"
Setup
reference.dirty = True
On Error Resume Next
guard.BeginGuard app, template
number = Err.Number
Err.Clear
On Error GoTo 0
Assert number <> 0, "Block pre-existing unsaved reference work"
Assert reference.setCalls = 0 And app.prefCalls = 0, "Preflight changes nothing"
guard.Restore
Setup
reference.failProtect = True
On Error Resume Next
guard.BeginGuard app, template
number = Err.Number
Err.Clear
On Error GoTo 0
Assert number <> 0, "Protection failure is visible"
guard.Restore
Assert app.preference = False, "Preference restored after partial BeginGuard"
Setup
guard.BeginGuard app, template
reference.failRestore = True
On Error Resume Next
guard.Restore
number = Err.Number
Err.Clear
On Error GoTo 0
Assert number <> 0, "Restoration failure is visible"
Assert app.preference = False, "Document restore failure cannot skip preference cleanup"
WScript.Echo "PASS: scoped protection, unsaved-work preflight, original states and partial-failure cleanup"
'''
        with tempfile.TemporaryDirectory(prefix="planter-save-guard-") as directory:
            runner = Path(directory) / "guard.vbs"
            runner.write_text(script, encoding="ascii")
            result = subprocess.run(["cscript", "//Nologo", str(runner)],
                                    capture_output=True, text=True, timeout=30)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertIn("PASS:", result.stdout, result.stdout + result.stderr)
