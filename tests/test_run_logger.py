"""Portable logger behavior with real filesystem writes, not a VBA compile test."""

from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
LOGGER = (ROOT / "src" / "UIRunLogger.cls").read_text(encoding="ascii")
MODULE = (ROOT / "src" / "userform_automation.bas").read_text(encoding="ascii")


class LoggingContracts(unittest.TestCase):
    def test_instrumentation_and_local_logs(self):
        self.assertIn("/run-logs/", (ROOT / ".gitignore").read_text())
        workflow = MODULE.split("Private Sub UIRunAutomation()", 1)[1]
        self.assertLess(workflow.index("runLogger.BeginRun"), workflow.index("tableSession.Apply"))
        self.assertIn('runLogger.Finish "FINISHED"', workflow)
        self.assertIn('runLogger.Finish "FAILED"', workflow)
        self.assertIn('UITraceBegin "Failure cleanup"', workflow)
        self.assertIn('UITraceDetail "Pack-and-Go document count"', workflow)
        self.assertIn('Record "INCOMPLETE"', LOGGER)
        self.assertIn('Record "SUMMARY"', LOGGER)
        self.assertIn("stream.Close", LOGGER)
        self.assertIn("enabled = False", LOGGER)
        self.assertIn("Runtime logging stopped", LOGGER)
        self.assertIn("fso.CreateTextFile(filePath, False, True)", LOGGER)
        self.assertNotIn("DeleteFile", LOGGER)

    def test_pack_and_go_profiling_contract(self):
        self.assertEqual(MODULE.count("template.Extension.SavePackAndGo("), 1)
        self.assertNotIn("packngo1.Step4_GetDocumentNames()", MODULE)
        for operation in ("Collection ForceRebuild3", "GetPackAndGo",
                          "Pack-and-Go IncludeDrawings", "Pack-and-Go GetDocumentNames",
                          "Source metadata inventory", "SavePackAndGo API"):
            self.assertIn(f'UITraceBegin "{operation}"', MODULE)
            self.assertIn(f'UITraceEnd "{operation}"', MODULE)
        self.assertLess(MODULE.index("    UITraceSourceFiles\n"),
                        MODULE.index("    If Not packngo1.Step6_BuildRenamedList()"))
        self.assertIn('UITraceDetail "Source metadata failure"', MODULE)
        self.assertIn("source-file metadata lookups failed", MODULE)
        self.assertIn('UITraceDetail "Opened packed assembly"', MODULE)


@unittest.skipUnless(shutil.which("cscript"), "Windows Script Host required")
class LoggerBehavior(unittest.TestCase):
    def test_collection_sequence_and_api_failures(self):
        body = re.search(r"^Private Sub UICollectPackAndGoDocuments\(\).*?^End Sub$",
                         MODULE, re.MULTILINE | re.DOTALL)[0]
        body = body.replace("Private Sub", "Sub").replace("packngo1.", "host.")
        # VBScript passes object properties by value; the VBA module global is ByRef.
        body = body.replace("host.docNames", "documentNames")
        script = r'''
Option Explicit
Dim host, events, result, number, description, documentNames
Const UI_ERROR = 1234
Set host = New FakeHost
Sub UITraceBegin(name)
    events = events & "BEGIN:" & name & "|"
End Sub
Sub UITraceEnd(name)
    events = events & "END:" & name & "|"
End Sub
Sub UITraceDetail(name, text)
End Sub
Sub Assert(condition, message)
    If Not condition Then
        WScript.Echo "FAIL: " & message
        WScript.Quit 1
    End If
End Sub
Class FakeHost
    Public swModel, swPackAndGo, usedManualFallback, docNames, docCount
    Private Sub Class_Initialize()
        Set swModel = New FakeModel
    End Sub
End Class
Class FakeModel
    Public Extension, rebuildResult
    Private Sub Class_Initialize()
        Set Extension = New FakeExtension
    End Sub
    Function ForceRebuild3(topOnly)
        events = events & "REBUILD|"
        ForceRebuild3 = rebuildResult
    End Function
End Class
Class FakeExtension
    Public pack, missing
    Private Sub Class_Initialize()
        Set pack = New FakePack
    End Sub
    Function GetPackAndGo()
        events = events & "GET_PACK|"
        If missing Then
            Set GetPackAndGo = Nothing
        Else
            Set GetPackAndGo = pack
        End If
    End Function
End Class
Class FakePack
    Public IncludeDrawings, IncludeSimulationResults, FlattenToSingleFolder, namesResult
    Function GetDocumentNames(names)
        events = events & "GET_NAMES|"
        names = Array("assembly.sldasm", "part.sldprt", "drawing.slddrw")
        GetDocumentNames = namesResult
    End Function
End Class
'''
        script += body + r'''
For Each result In Array(1, -1)
    events = ""
    host.swModel.rebuildResult = result
    host.swModel.Extension.pack.namesResult = result
    UICollectPackAndGoDocuments
    Assert host.docCount = 3, "Document count"
    Assert host.swPackAndGo.IncludeDrawings, "Drawings retained"
    Assert Not host.swPackAndGo.IncludeSimulationResults, "Simulation excluded"
    Assert host.swPackAndGo.FlattenToSingleFolder, "Flattening retained"
    Assert Not host.usedManualFallback, "No unsafe fallback"
    Assert InStr(events, "REBUILD|") < InStr(events, "GET_PACK|"), "Rebuild before collection"
    Assert InStr(events, "GET_PACK|") < InStr(events, "GET_NAMES|"), "Object before discovery"
Next
host.swModel.rebuildResult = 0
events = ""
On Error Resume Next
UICollectPackAndGoDocuments
number = Err.Number
Err.Clear
On Error GoTo 0
Assert number = UI_ERROR, "Failed rebuild stops"
Assert InStr(events, "GET_PACK|") = 0, "No collection after failed rebuild"
host.swModel.rebuildResult = 1
host.swModel.Extension.missing = True
On Error Resume Next
UICollectPackAndGoDocuments
number = Err.Number
Err.Clear
On Error GoTo 0
Assert number = UI_ERROR, "Missing object stops"
host.swModel.Extension.missing = False
host.swModel.Extension.pack.namesResult = 0
On Error Resume Next
UICollectPackAndGoDocuments
number = Err.Number
description = Err.Description
Err.Clear
On Error GoTo 0
Assert number = UI_ERROR, "Failed native discovery stops"
Assert InStr(description, "manual fallback") > 0, "Explicit unsafe fallback error"
WScript.Echo "PASS: collection sequence, options, True=1/-1, failed rebuild/object/discovery"
'''
        with tempfile.TemporaryDirectory(prefix="planter-collection-test-") as directory:
            runner = Path(directory) / "collection.vbs"
            runner.write_text(script, encoding="ascii")
            result = subprocess.run(["cscript", "//Nologo", str(runner)],
                                    capture_output=True, text=True, timeout=30)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertIn("PASS:", result.stdout, result.stdout + result.stderr)

    def test_files_timings_errors_and_collision(self):
        body = LOGGER.split("Option Explicit", 1)[1]
        body = re.sub(r" As (?:String|Object|Double|Boolean|Long|Variant)\b", "", body)
        body = re.sub(r"^(\s*Next) \w+$", r"\1", body, flags=re.MULTILINE)
        body = body.replace("Format$", "TestFormat")
        body = body.replace("Dim stream\n", "Dim stream\n    Set stream = Nothing\n")
        body = body.replace("On Error GoTo Failed", "On Error Resume Next\n    Err.Clear")
        body = body.replace("    Exit Sub\nFailed:", "    If Err.Number = 0 Then Exit Sub")
        script = r'''
Option Explicit
Dim clock, warnings, folder, logger, second, path1, path2, fso
clock = 1000
warnings = 0
folder = WScript.Arguments(0)
Set fso = CreateObject("Scripting.FileSystemObject")
Function UIRunClock()
    UIRunClock = clock
End Function
Function UIInchesText(value)
    UIInchesText = Trim(Str(value))
End Function
Function Str(value)
    Str = CStr(value)
End Function
Function TestFormat(value, pattern)
    TestFormat = "timestamp"
End Function
Sub MsgBox(text, style, title)
    warnings = warnings + 1
End Sub
Sub Assert(condition, message)
    If Not condition Then
        WScript.Echo "FAIL: " & message
        WScript.Quit 1
    End If
End Sub
Class UIRunLogger
'''
        script += body + r'''
End Class
Set logger = New UIRunLogger
logger.BeginRun folder
path1 = logger.Path
logger.Detail "Input", "a" & vbTab & "b" & vbCrLf & "c"
logger.BeginSpan "Stage one"
clock = 1002.5
logger.EndSpan "Stage one"
logger.BeginSpan "Stage two"
clock = 1005
logger.Finish "FAILED", "test failure"
Set second = New UIRunLogger
second.BeginRun folder
path2 = second.Path
Assert path1 <> path2, "Never overwrite a same-second log"
second.BeginSpan "Completed stage"
clock = 1008
second.EndSpan "Completed stage"
second.Finish "FINISHED", "done"
Assert warnings = 0, "Normal logging must not warn"
Set logger = New UIRunLogger
logger.BeginRun folder
fso.DeleteFile logger.Path
logger.Detail "Write failure", "test"
Assert warnings = 1, "Runtime write failure warns explicitly"
logger.Detail "Disabled", "No repeated warnings"
Assert warnings = 1, "Failed logging disables subsequent writes"
Set logger = New UIRunLogger
On Error Resume Next
logger.BeginRun folder & "\missing-parent\logs"
Assert Err.Number <> 0, "Invalid initial log folder propagates an error before editing models"
Err.Clear
On Error GoTo 0
WScript.Echo "PASS: files, spans, incomplete stages, summaries, collisions and write-error handling"
'''
        with tempfile.TemporaryDirectory(prefix="planter-log-test-") as directory:
            runner = Path(directory) / "logger.vbs"
            folder = Path(directory) / "logs"
            runner.write_text(script, encoding="ascii")
            result = subprocess.run(["cscript", "//Nologo", str(runner), str(folder)],
                                    capture_output=True, text=True, timeout=30)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertIn("PASS:", result.stdout, result.stdout + result.stderr)
            logs = list(folder.glob("*.log"))
            self.assertEqual(len(logs), 2)
            failed = (folder / "Run_timestamp.log").read_text(encoding="utf-16")
            self.assertIn("\tEND\tStage one\t2.5\t", failed)
            self.assertIn("\tINCOMPLETE\tStage two\t2.5\tFAILED", failed)
            self.assertIn("\tSUMMARY\tStage one\t2.5\t", failed)
            self.assertIn("\tRUN_FAILED\tRunWithUI\t5\t", failed)
            self.assertIn("\tDETAIL\tInput\t\ta b  c", failed)
            finished = (folder / "Run_timestamp_1.log").read_text(encoding="utf-16")
            self.assertIn("\tRUN_FINISHED\tRunWithUI\t3\t", finished)
