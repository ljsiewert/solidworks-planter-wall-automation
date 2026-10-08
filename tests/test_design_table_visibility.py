"""Workbook-scoped visibility cleanup; fake-host execution, not Excel integration."""
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
SOURCE = (ROOT / "src" / "UIDesignTableSession.cls").read_text(encoding="ascii")


class DesignTableVisibilityTests(unittest.TestCase):
    def test_visibility_is_workbook_scoped_and_restored_before_close(self):
        self.assertIn("For Each window In sheet.Parent.Windows", SOURCE)
        self.assertNotIn("Application.Visible", SOURCE)
        self.assertNotIn("ScreenUpdating", SOURCE)
        self.assertNotIn("ActiveWindow", SOURCE)
        opening = SOURCE.split("Private Sub OpenSheet()", 1)[1].split("End Sub", 1)[0]
        self.assertIn("HideWorkbookWindows", opening)
        closing = SOURCE.split("Private Sub CloseSheet(", 1)[1].split("End Sub", 1)[0]
        self.assertLess(closing.index("RestoreWorkbookWindows"), closing.index("Set sheet = Nothing"))
        self.assertLess(closing.index("Set sheet = Nothing"), closing.index("table.UpdateTable"))
        cleanup = SOURCE.split("Public Sub CloseEditor()", 1)[1].split("End Sub", 1)[0]
        self.assertLess(cleanup.index("RestoreWorkbookWindows"), cleanup.index("If table Is Nothing"))

    @unittest.skipUnless(shutil.which("cscript"), "Windows Script Host required")
    def test_snapshot_normal_close_and_failure_cleanup(self):
        methods = []
        for name in ("HideWorkbookWindows", "RestoreWorkbookWindows", "CloseSheet", "CloseEditor"):
            body = re.search(rf"^(?:Private|Public) Sub {name}\b.*?^End Sub$",
                             SOURCE, re.MULTILINE | re.DOTALL)[0]
            body = re.sub(r" As (?:Object|Long)\b", "", body)
            body = body.replace("Private Sub", "Public Sub").replace("New Collection", "New FakeCollection")
            body = body.replace("swDesignTableUpdateOptions_e.swUpdateDesignTableNone", "3")
            body = re.sub(r"^(\s*Next) \w+$", r"\1", body, flags=re.MULTILINE)
            methods.append(body)
        script = r'''
Option Explicit
Const TABLE_ERROR = 2200
Dim session, first, second, unrelated, workbook, worksheet, table, errorNumber
Class FakeCollection
    Private items
    Private Sub Class_Initialize()
        Set items = CreateObject("Scripting.Dictionary")
    End Sub
    Public Sub Add(value)
        items.Add CStr(items.Count + 1), value
    End Sub
    Public Sub Remove(index)
        items.Remove CStr(index)
    End Sub
    Public Property Get Count()
        Count = items.Count
    End Property
    Public Default Property Get Item(index)
        If IsObject(items(CStr(index))) Then
            Set Item = items(CStr(index))
        Else
            Item = items(CStr(index))
        End If
    End Property
End Class
Class FakeWindow
    Private shown
    Public failHide, failRestore
    Public Property Get Visible()
        Visible = shown
    End Property
    Public Property Let Visible(value)
        If (Not value And failHide) Or (value And failRestore) Then
            Err.Raise 2201, , "Visibility failure"
        End If
        shown = value
    End Property
End Class
Class FakeWorkbook
    Public Windows
End Class
Class FakeSheet
    Public Parent
End Class
Class FakeTable
    Public IsActive, updates, lastOption, failUpdate
    Public Function UpdateTable(updateOption, closeWindow)
        Assert first.Visible And Not second.Visible, "Original visibility restored before commit"
        Assert session.sheet Is Nothing, "Worksheet released before commit"
        Assert session.workbookWindows Is Nothing, "Window references released before commit"
        updates = updates + 1
        lastOption = updateOption
        If failUpdate Then
            UpdateTable = False
        Else
            IsActive = False
            UpdateTable = 1
        End If
    End Function
End Class
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
Sub Setup()
    Set first = New FakeWindow
    first.Visible = True
    Set second = New FakeWindow
    second.Visible = False
    Set unrelated = New FakeWindow
    unrelated.Visible = True
    Set workbook = New FakeWorkbook
    workbook.Windows = Array(first, second)
    Set worksheet = New FakeSheet
    Set worksheet.Parent = workbook
    Set table = New FakeTable
    table.IsActive = True
    table.updates = 0
    Set session = New FakeSession
    Set session.saveGuard = Nothing
    Set session.workbookWindows = Nothing
    Set session.windowVisibility = Nothing
    Set session.sheet = worksheet
    Set session.table = table
End Sub
Class FakeSession
    Public sheet, table, saveGuard, workbookWindows, windowVisibility
'''
        script += "\n".join(methods) + r'''
End Class
Setup
session.HideWorkbookWindows
Assert Not first.Visible And Not second.Visible, "Only template windows hidden"
Assert unrelated.Visible, "Other workbook untouched"
session.CloseSheet 2
Assert table.updates = 1 And table.lastOption = 2, "Normal update behavior preserved"
session.CloseEditor
Assert table.updates = 1, "Cleanup after success is idempotent"

Setup
second.failHide = True
On Error Resume Next
session.HideWorkbookWindows
errorNumber = Err.Number
On Error GoTo 0
Assert errorNumber = 2201, "Hide error propagates"
second.failHide = False
session.CloseEditor
Assert first.Visible And Not second.Visible And unrelated.Visible, "Partial hide restored on failure"
Assert table.lastOption = 3, "Failed editing closes without applying table"

Setup
session.HideWorkbookWindows
first.failRestore = True
On Error Resume Next
session.CloseEditor
errorNumber = Err.Number
On Error GoTo 0
Assert errorNumber = 2201 And table.updates = 0, "Restore errors propagate before closing editor"
Assert session.workbookWindows.Count = 1, "Failed restoration snapshot retained for retry"
first.failRestore = False
session.CloseEditor
Assert first.Visible And table.updates = 1, "Cleanup retry restores outstanding window"

Setup
session.HideWorkbookWindows
table.failUpdate = True
On Error Resume Next
session.CloseSheet 2
errorNumber = Err.Number
On Error GoTo 0
Assert errorNumber = TABLE_ERROR, "Commit failure reported"
Assert first.Visible And unrelated.Visible, "Editor visible for review after commit failure"
table.failUpdate = False
session.CloseEditor
Assert table.updates = 2, "Failed commit can still close editor"
WScript.Echo "PASS: scoped visibility and failure cleanup"
'''
        with tempfile.TemporaryDirectory(prefix="planter-visibility-") as folder:
            runner = Path(folder) / "visibility.vbs"
            runner.write_text(script, encoding="ascii")
            result = subprocess.run(["cscript", "//Nologo", str(runner)], capture_output=True,
                                    text=True, timeout=30)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertIn("PASS:", result.stdout, result.stdout + result.stderr)
