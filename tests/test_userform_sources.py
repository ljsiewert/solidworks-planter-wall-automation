"""Source contracts and portable helper tests; not a SolidWorks/VBA compile test.

On Windows, the actual VBA validation/mapping function bodies are also executed
by cscript after stripping type annotations. VBA Val/IsError are supplied for the
already-validated decimal grammar, with the test process locale set to en-US.
"""

from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import unittest


ROOT = Path(__file__).resolve().parents[1]
MODULE = (ROOT / "src" / "userform_automation.bas").read_text(encoding="ascii")
FORM = (ROOT / "src" / "UserForm_AutomationUI.frm").read_text(encoding="ascii")
PACKNGO = (ROOT / "src" / "packngo.bas").read_text(encoding="utf-8")
TABLE = (ROOT / "src" / "UIDesignTableSession.cls").read_text(encoding="ascii")


def helper_body(name, source=MODULE):
    match = re.search(
        rf"^(?:Public )?(Function|Sub) {name}\b.*?^End \1$",
        source,
        flags=re.MULTILINE | re.DOTALL,
    )
    if not match:
        raise AssertionError(f"Missing helper: {name}")
    body = re.sub(r" As (?:String|Double|Long|Boolean|Object|Variant)\b", "", match[0])
    body = re.sub(r"^(\s*Next) \w+$", r"\1", body, flags=re.MULTILINE)
    return re.sub(r"\b(Trim|Mid|Right|Left)\$", r"\1", body)


class SourceContracts(unittest.TestCase):
    def test_dual_entry_point_and_safe_workflow(self):
        self.assertIn("Public Sub RunWithUI()", MODULE)
        self.assertIn("Set form = New UserForm_AutomationUI", MODULE)
        self.assertIn("form.Show vbModal", MODULE)
        self.assertLess(MODULE.index("If form.Accepted Then"), MODULE.index("        UIRunAutomation"))
        for forbidden in (
            "packngo1.RunPackAndGo",
            "packngo1.Step5_SetDestination",
            "packngo1.Step8_ConfirmAndSave",
            "packngo1.PromptAndRunExports",
            "InputBox",
            "On Error Resume Next",
            "DeleteFolder",
            "DeleteFile",
        ):
            self.assertNotIn(forbidden, MODULE)
        for helper in ("Step3_BuildNames", "Step4_GetDocumentNames",
                       "Step6_BuildRenamedList", "Step7_ApplyRenamedList"):
            self.assertIn("packngo1." + helper, MODULE)

    def test_design_table_and_template_safety(self):
        for equation in ("WALL_COUNT", "SIDE", "RETURN_TYPE", "OVERALL_LENGTH",
                         "OVERALL_WIDTH", "OVERALL_HEIGHT", "CENTER_PLANTER_LENGTH",
                         "SIDE_PLANTER_LENGTH", "OVERALL_MATERIAL", "OVERALL_THICKNESS"):
            self.assertIn(f'"{equation}"', MODULE)
        self.assertNotIn("SetEquationAndConfigurationOption", MODULE)
        self.assertNotRegex(MODULE, r"\.Equation\([^)]*\)\s*=")
        self.assertIn("tableSession.Apply template, designTableInputs, expected", MODULE)
        self.assertIn("UIReadCalculatedGlobals template, expected", MODULE)
        self.assertIn("tableSession.Restore", MODULE)
        self.assertIn("UIVerifyDesignTableModel packed, expected", MODULE)
        self.assertIn("configurationName, errors, warnings", MODULE)
        self.assertNotIn("template.Save", MODULE)
        self.assertNotIn("CloseDoc", MODULE)
        self.assertIn('originals = sheet.Range("N3:N9").Formula', TABLE)
        self.assertIn('sheet.Range("N3:N9").Formula = originals', TABLE)
        self.assertIn('sheet.Cells(i + 3, 14).Value2 = inputs(i)', TABLE)
        self.assertIn('sheet.Range("N8").NumberFormat = "@"', TABLE)
        self.assertIn("sheet.Calculate", TABLE)
        self.assertIn("VerifyFormulas", TABLE)
        self.assertIn("If table.LinkToFile Then", TABLE)
        self.assertIn("If table.IsActive Then", TABLE)
        self.assertIn('sheet.Range("A3").Value2', TABLE)
        self.assertIn("swUpdateDesignTableAll", TABLE)
        self.assertNotIn("N12", TABLE)
        self.assertNotIn(".Quit", TABLE)
        self.assertIn("If Not changed Then", TABLE)
        self.assertIn("VerifyModel originalOutputs", TABLE)
        self.assertLess(TABLE.index("Set sheet = Nothing"), TABLE.index("table.UpdateTable"))

    def test_solidworks_boolean_checks_do_not_use_bitwise_not(self):
        for source in (MODULE, TABLE):
            self.assertNotRegex(
                source,
                r"If Not [^\r\n]*(?:ForceRebuild3|ShowConfiguration2|UpdateTable|Save3|SetSaveToName|RunExternalMacro)\(",
            )
        for expression in (
            "model.ShowConfiguration2(configurationName)",
            "model.ForceRebuild3(False)",
            "table.UpdateTable(updateOption, True)",
        ):
            self.assertIn(f"If {expression} = False Then", TABLE)
        self.assertIn("If packed.ForceRebuild3(False) = False Then", MODULE)

    def test_export_routes_and_quote_only(self):
        for file, procedure in (
            ("pdfassemblies.swp", "RunPDFAssemblyDrawings"),
            ("pdfcomponents.swp", "RunPDFComponentDrawings"),
            ("dxfcomponents.swp", "RunDXFComponentDrawings"),
            ("stepcomponents.swp", "RunSTEPComponentExport"),
        ):
            self.assertIn(f'"{file}", "{procedure}"', MODULE)
        for flag, control in (
            ("exportPDFComponents", "pdfComponentsBox"),
            ("exportDXF", "dxfBox"),
            ("exportSTEP", "stepBox"),
        ):
            self.assertIn(f"{flag} = Not quoteOnly And CBool({control}.Value)", FORM)
            self.assertIn(f"{control}.Value = False", FORM)
        self.assertIn("exportPDFAssemblies = quoteOnly Or CBool(pdfAssembliesBox.Value)", FORM)
        self.assertIn("pdfAssembliesBox.Value = True", FORM)
        self.assertIn("UICheckExportFiles", FORM)

    def test_form_controls_and_events(self):
        for name in ("cboWallCount", "cboSide", "cboReturnType", "txtOverallLength",
                     "txtOverallWidth", "txtOverallHeight", "txtCenterLength",
                     "txtSideLength", "cboMaterial", "cboThickness", "txtOutputFolder",
                     "chkQuoteOnly", "chkPDFAssemblies", "chkPDFComponents", "chkDXF",
                     "chkSTEP", "cmdRun", "cmdCancel", "cmdBrowse"):
            self.assertIn(f'"{name}"', FORM)
        for event in ("runButton_Click", "cancelButton_Click", "browseButton_Click",
                      "wallsBox_Change", "quoteBox_Click"):
            self.assertIn(f"Private Sub {event}()", FORM)
            member = event.split("_")[0]
            self.assertIn(f"Private WithEvents {member} As MSForms.", FORM)
        self.assertIn("Private Sub UserForm_QueryClose", FORM)
        self.assertIn("combo.Style = fmStyleDropDownList", FORM)
        self.assertIn("acceptedValue = False", FORM)
        self.assertIn("shell.BrowseForFolder", FORM)

    def test_dimensions_and_calculated_fields(self):
        self.assertIn("UIInchesText = Trim$(Str$(value))", MODULE)
        self.assertIn("CStr(length)", FORM)
        self.assertIn("UIValidateDimensions length, width, height", FORM)
        self.assertIn("centerBox.Locked = True", FORM)
        self.assertIn("sideLengthBox.Locked = True", FORM)
        self.assertNotIn("ReadDimension(centerBox", FORM)
        self.assertNotIn("ReadDimension(sideLengthBox", FORM)
        self.assertIn("designTableInputs = Array(", FORM)


@unittest.skipUnless(shutil.which("cscript"), "Windows Script Host is required")
class PortableHelperTests(unittest.TestCase):
    def test_rebuild_configuration_activation(self):
        body = re.search(
            r"^Private Sub Rebuild\b.*?^End Sub$", TABLE,
            flags=re.MULTILINE | re.DOTALL,
        ).group(0)
        body = body.replace("Private Sub", "Sub", 1)
        body = re.sub(r" As String\b", "", body)
        script = r'''
Option Explicit
Const TABLE_ERROR = -2147219304
Dim model, configurationName, number, representation
Class FakeConfiguration
    Public Name
End Class
Class FakeManager
    Public ActiveConfiguration
    Private Sub Class_Initialize()
        Set ActiveConfiguration = New FakeConfiguration
    End Sub
End Class
Class FakeModel
    Public ConfigurationManager, activateCalls, rebuildCalls
    Public activationResult, changeName, rebuildResult
    Private Sub Class_Initialize()
        Set ConfigurationManager = New FakeManager
        activateCalls = 0
        rebuildCalls = 0
        activationResult = True
        changeName = True
        rebuildResult = True
    End Sub
    Public Function ShowConfiguration2(name)
        activateCalls = activateCalls + 1
        If activationResult <> 0 And changeName Then ConfigurationManager.ActiveConfiguration.Name = name
        ShowConfiguration2 = activationResult
    End Function
    Public Function ForceRebuild3(topOnly)
        rebuildCalls = rebuildCalls + 1
        ForceRebuild3 = rebuildResult
    End Function
End Class
Sub Assert(condition, message)
    If Not condition Then
        WScript.Echo "FAIL: " & message
        WScript.Quit 1
    End If
End Sub
Sub ExpectFailure()
    Dim number
    On Error Resume Next
    Rebuild "test"
    number = Err.Number
    Err.Clear
    On Error GoTo 0
    Assert number <> 0, "Expected a visible configuration/rebuild failure"
End Sub
'''
        script += "\n" + body + r'''
configurationName = "PLANTER ASSEMBLY"
Set model = New FakeModel
model.ConfigurationManager.ActiveConfiguration.Name = "PLANTER ASSEMBLY"
model.activationResult = False
Rebuild "test"
Assert model.activateCalls = 0, "Already-active configuration must not be reactivated"
Assert model.rebuildCalls = 1, "Already-active configuration must still rebuild"
Set model = New FakeModel
model.ConfigurationManager.ActiveConfiguration.Name = "planter assembly"
Rebuild "test"
Assert model.activateCalls = 0, "Active configuration comparison is case insensitive"
Set model = New FakeModel
model.ConfigurationManager.ActiveConfiguration.Name = "Other"
Rebuild "test"
Assert model.activateCalls = 1 And model.rebuildCalls = 1, "Switch then rebuild"
Set model = New FakeModel
model.ConfigurationManager.ActiveConfiguration.Name = "Other"
model.activationResult = False
ExpectFailure
Assert model.rebuildCalls = 0, "Do not rebuild after a failed switch"
Set model = New FakeModel
model.ConfigurationManager.ActiveConfiguration.Name = "Other"
model.changeName = False
ExpectFailure
Assert model.rebuildCalls = 0, "Check actual active name, not only API return"
Set model = New FakeModel
model.ConfigurationManager.ActiveConfiguration.Name = configurationName
model.rebuildResult = False
ExpectFailure
Assert model.rebuildCalls = 1, "Rebuild failures must still be reported"
For Each representation In Array(1, -1)
    Set model = New FakeModel
    model.ConfigurationManager.ActiveConfiguration.Name = configurationName
    model.rebuildResult = representation
    Rebuild "test"
    Assert model.rebuildCalls = 1, "Accept both nonzero representations of True"
    Set model = New FakeModel
    model.ConfigurationManager.ActiveConfiguration.Name = "Other"
    model.activationResult = representation
    model.rebuildResult = representation
    Rebuild "test"
    Assert model.activateCalls = 1 And model.rebuildCalls = 1, "Accept True=1 for configuration switch"
Next
WScript.Echo "PASS: already-active, switched, failed-switch and failed-rebuild paths"
'''
        with tempfile.TemporaryDirectory(prefix="planter-ui-config-test-") as directory:
            runner = Path(directory) / "rebuild.vbs"
            runner.write_text(script, encoding="ascii")
            result = subprocess.run(
                ["cscript", "//Nologo", str(runner)], capture_output=True,
                text=True, check=False, timeout=30,
            )
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertIn("PASS:", result.stdout, result.stdout + result.stderr)

    def test_validation_mappings_and_output_paths(self):
        helpers = "\n\n".join(helper_body(name) for name in (
            "UIParseInches", "UIValidateSelections", "UIMapSelections", "UIOutputRoot",
            "UIValidateDimensions", "UIWallTableLabel", "UIReturnTableLabel", "UIReadTableNumber",
            "UIProgressTime",
        ))
        helpers += "\n\n" + "\n\n".join(helper_body(name, PACKNGO) for name in (
            "Step3_BuildNames", "FormatPNDimension",
        ))
        script = r'''
Option Explicit
Const UI_ERROR = -2147219404
Dim packngo1, walls, orientation, returnIndex, materialIndex, thicknessIndex
Dim returns, materials, thicknesses, orientations, good, bad, item, count, root
Dim fso, expectedSide, actual, ignored
Dim wallCount, side, returnType, overallMaterial, thicknessLabel
Dim overallLength, overallWidth, overallHeight, centerPlanterLength, sidePlanterLength
Dim masterName, centerName, centerHardwareName, centerGussetName, mountingGussetName
Dim sideName, sideLeftName, sideRightName, sideCenterName, sideWallName, sharedSideName
SetLocale 1033
Class PackState
    Public wallCount, side, returnType, overallMaterial, thicknessLabel
End Class
Set packngo1 = New PackState
Function Val(text)
    Val = CDbl(text)
End Function
Function IsError(value)
    IsError = (VarType(value) = 10)
End Function
Sub Assert(condition, message)
    If Not condition Then
        WScript.Echo "FAIL: " & message
        WScript.Quit 1
    End If
End Sub
Sub ExpectBadDimension(text)
    Dim value, number
    On Error Resume Next
    value = UIParseInches(text)
    number = Err.Number
    Err.Clear
    On Error GoTo 0
    Assert number <> 0, "Accepted invalid measurement: [" & text & "]"
End Sub
Sub ExpectBadSelection(walls, orientation, returnLabel, material, thickness)
    Dim number
    On Error Resume Next
    UIValidateSelections walls, orientation, returnLabel, material, thickness
    number = Err.Number
    Err.Clear
    On Error GoTo 0
    Assert number <> 0, "Accepted an unsupported selection"
End Sub
Sub ExpectBadDimensions(length, width, height)
    Dim number
    On Error Resume Next
    UIValidateDimensions length, width, height
    number = Err.Number
    Err.Clear
    On Error GoTo 0
    Assert number <> 0, "Accepted dimensions outside workbook validation"
End Sub
'''
        script += "\n" + helpers + r'''
good = Array("1", "100.25", ".25", "1.", "0001.25", " 10.50 ", "100000", "100001", "0.0001")
For Each item In good
    Assert UIParseInches(item) = CDbl(Trim(item)), "Wrong measurement: " & item
Next
bad = Array("", " ", ".", "0", "0.0", "-1", "+1", "1/4", "1 in", "1,25", _
            "1e2", "1.2.3", "NaN", "1 2", "1" & vbTab)
For Each item In bad
    ExpectBadDimension item
Next
returns = Array("No return", "Single return 1 inch", "Single return 2 inches", "Double return")
materials = Array("Mild Steel", "Borcon Weathering Steel")
thicknesses = Array("3/16 inch", "1/4 inch")
count = 0
For walls = 1 To 4
    If walls = 2 Then
        orientations = Array("Left", "Right")
    Else
        orientations = Array("N/A")
    End If
    For Each orientation In orientations
        For returnIndex = 0 To 3
            For materialIndex = 0 To 1
                For thicknessIndex = 0 To 1
                    UIMapSelections walls, orientation, returns(returnIndex), _
                                    materials(materialIndex), thicknesses(thicknessIndex)
                    expectedSide = "0"
                    If orientation = "Left" Then expectedSide = "2"
                    If orientation = "Right" Then expectedSide = "1"
                    Assert packngo1.wallCount = CStr(walls), "Wrong wall count"
                    Assert packngo1.side = expectedSide, "Wrong side code"
                    Assert packngo1.returnType = CStr(returnIndex), "Wrong return code"
                    Assert packngo1.overallMaterial = CStr(materialIndex + 1), "Wrong material code"
                    Assert packngo1.thicknessLabel = CStr(thicknessIndex + 1), "Wrong thickness code"
                    count = count + 1
                Next
            Next
        Next
    Next
Next
Assert count = 80, "Expected all 80 supported selection combinations"
Assert UIProgressTime(0) = "0m 0s", "Elapsed time at start"
Assert UIProgressTime(59) = "0m 59s", "Sub-minute elapsed time"
Assert UIProgressTime(60) = "1m 0s", "Minute rollover"
Assert UIProgressTime(3661) = "61m 1s", "Long-running duration"
Assert UIWallTableLabel(1, "N/A") = "ONE WALL", "N3 one wall"
Assert UIWallTableLabel(2, "Left") = "TWO WALLS (CENTER AND LEFT SIDE)", "N3 left"
Assert UIWallTableLabel(2, "Right") = "TWO WALLS (CENTER AND RIGHT SIDE)", "N3 right"
Assert UIWallTableLabel(3, "N/A") = "THREE WALLS", "N3 three"
Assert UIWallTableLabel(4, "N/A") = "FOUR WALLS", "N3 four"
Assert UIReturnTableLabel(returns(0)) = "Default", "N4 default"
Assert UIReturnTableLabel(returns(1)) = "single_return1", "N4 single1"
Assert UIReturnTableLabel(returns(2)) = "single_return2", "N4 single2"
Assert UIReturnTableLabel(returns(3)) = "double_return", "N4 double"
Assert UIReadTableNumber("=113", "N10") = 113, "Calculated string value"
Assert UIReadTableNumber("=100.4", "N11") = 100.4, "Calculated side value"
Assert UIReadTableNumber("=0", "C3") = 0, "Zero orientation"
Assert UIReadTableNumber(0.1875, "K3") = 0.1875, "Numeric thickness"
UIValidateDimensions 60, 60, 6
UIValidateDimensions 60, 60, 48
ExpectBadDimensions 60, 60, 5.999
ExpectBadDimensions 60, 60, 48.001
ExpectBadDimensions 48, 60, 48
ExpectBadDimensions 60, 48, 48
ExpectBadDimensions 6, 60, 6
UIMapSelections 4, "N/A", returns(0), materials(0), thicknesses(0)
wallCount = packngo1.wallCount
side = packngo1.side
returnType = packngo1.returnType
overallMaterial = packngo1.overallMaterial
thicknessLabel = packngo1.thicknessLabel
overallLength = "100"
overallWidth = "100"
overallHeight = "10"
centerPlanterLength = "100.25"
sidePlanterLength = "10.50"
Assert Step3_BuildNames(), "Existing naming helper failed"
Assert masterName = "40-0-C10000-S10000-1000-11", "Main assembly PDF numbering"
Assert centerName = "0-C10025-1000-11", "Center assembly PDF numbering"
Assert sideLeftName = "0-SL1050-1000-11", "Left assembly PDF numbering"
Assert centerGussetName = "1000-11-CG", "Gusset PDF numbering"
Assert FormatPNDimension(10.254) = "1025", "Existing dimension rounding below half"
Assert FormatPNDimension(10.256) = "1026", "Existing dimension rounding above half"
ExpectBadSelection 0, "N/A", returns(0), materials(0), thicknesses(0)
ExpectBadSelection 5, "N/A", returns(0), materials(0), thicknesses(0)
ExpectBadSelection 2, "Both", returns(0), materials(0), thicknesses(0)
ExpectBadSelection 2, "N/A", returns(0), materials(0), thicknesses(0)
ExpectBadSelection 3, "Left", returns(0), materials(0), thicknesses(0)
ExpectBadSelection 1, "N/A", "Open", materials(0), thicknesses(0)
ExpectBadSelection 1, "N/A", returns(0), "Aluminum", thicknesses(0)
ExpectBadSelection 1, "N/A", returns(0), materials(0), "1/8 inch"
Set fso = CreateObject("Scripting.FileSystemObject")
root = WScript.Arguments(0)
Assert UIOutputRoot(root) = root & "\PACK_TEST\", "Output root should gain PACK_TEST"
Assert UIOutputRoot(root & "\PACK_TEST") = root & "\PACK_TEST\", "Do not duplicate PACK_TEST"
Assert LCase(UIOutputRoot(root & "\pack_test")) = LCase(root & "\pack_test\"), "Case insensitive marker"
Assert UIOutputRoot(root & "\not_PACK_TEST") = root & "\not_PACK_TEST\PACK_TEST\", "Use whole path segments"
On Error Resume Next
ignored = UIOutputRoot(root & "\missing")
actual = Err.Number
Err.Clear
On Error GoTo 0
Assert actual <> 0, "Missing output root accepted"
WScript.Echo "PASS: decimals, 80 PDF mappings, existing naming, invalid selections, output paths"
'''
        with tempfile.TemporaryDirectory(prefix="planter-ui-test-") as directory:
            folder = Path(directory)
            (folder / "PACK_TEST").mkdir()
            (folder / "not_PACK_TEST").mkdir()
            runner = folder / "helpers.vbs"
            runner.write_text(script, encoding="ascii")
            result = subprocess.run(
                ["cscript", "//Nologo", str(runner), str(folder)],
                capture_output=True, text=True, check=False, timeout=30,
            )
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertIn("PASS:", result.stdout, result.stdout + result.stderr)


if __name__ == "__main__":
    unittest.main()
