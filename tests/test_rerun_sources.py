"""Rerun-from-output and exports-only contracts, plus a cscript round trip of the
packed-name -> template-role mapping using the real packngo naming helpers."""

from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import unittest

from test_userform_sources import FORM, MODULE, PACKNGO, ROOT, helper_body


RERUN = (ROOT / "src" / "ui_rerun.bas").read_text(encoding="ascii")

TEMPLATE_FILES = [
    "planter_assembly.sldasm", "planter_assembly.slddrw",
    "planter_center_asm.sldasm", "planter_center_asm.slddrw",
    "planterwallcenter.sldprt", "planterwallcenter.slddrw",
    "baseplatecenter.sldprt", "baseplatecenter.slddrw",
    "centergusset.sldprt", "centergusset.slddrw",
    "mountinggusset.sldprt", "mountinggusset.slddrw",
    "planter_side_asm.sldasm", "planter_side_asm.slddrw",
    "planter_side_asm_left.sldasm", "planter_side_asm_left.slddrw",
    "planter_side_asm_right.sldasm", "planter_side_asm_right.slddrw",
    "planter_side_asm_center.sldasm", "planter_side_asm_center.slddrw",
    "planterwallside.sldprt", "planterwallside.slddrw",
    "planterwallsideleft.sldprt", "planterwallsideleft.slddrw",
    "planterwallsideright.sldprt", "planterwallsideright.slddrw",
    "planterwallsidecenter.sldprt", "planterwallsidecenter.slddrw",
    "baseplateside.sldprt", "baseplateside.slddrw",
    "leftsidewall.sldprt", "leftsidewall.slddrw",
    "rightsidewall.sldprt", "rightsidewall.slddrw",
]


class RerunContracts(unittest.TestCase):
    def test_equation_errors_are_reported_without_whats_wrong_items(self):
        body = helper_body("UIRerunWhatsWrongOf", RERUN.replace("Private Sub", "Sub"))
        self.assertLess(body.index("equationText = manager.Equation(i)"),
                        body.index("equationStatus = manager.Status"))
        self.assertLess(body.index("If equationStatus < 0 Then"),
                        body.index("If document.Extension.GetWhatsWrongCount = 0 Then Exit Sub"))
        self.assertIn("equationText", body[body.index("If equationStatus < 0 Then"):])
        self.assertIn("lines = lines + 1", body)

    def test_entry_routes_to_rerun_and_exports_only(self):
        run = helper_body("RunWithUI")
        self.assertLess(run.index("UIDetectRerunSource"), run.index("Set form = New UserForm_AutomationUI"))
        self.assertLess(run.index("If Len(rerunPath) > 0 Then"), run.index("planter_assembly.SLDASM"))
        self.assertIn("UIRunExportsOnly", run)
        self.assertIn("UIRunRerun", run)
        for name in ("UIRunExport", "UIReadCalculatedGlobals", "UICloseTableEditor"):
            self.assertRegex(MODULE, rf"Public (Sub|Function) {name}\(")

    def test_rerun_module_is_safe(self):
        self.assertIn('Attribute VB_Name = "ui_rerun"', RERUN)
        for forbidden in ("On Error Resume Next", "DeleteFile", "DeleteFolder", "InputBox"):
            self.assertNotIn(forbidden, RERUN)
        rerun = helper_body("UIRunRerun", RERUN)
        order = ["UIRerunSaveOpenWork", "UIRerunCloseOutput app, sourcePath", "UIRerunCreateStage",
                 "Set source = app.OpenDoc6(stagedPath", "tableSession.Apply source", "UIRerunBuildNames",
                 "GetPackAndGo", "GetDocumentNames",
                 "SetDocumentSaveToNames", "SavePackAndGo", "UIRerunCloseOutput app, stagedPath",
                 "UIRerunCleanStage", "Set packed = app.OpenDoc6", "tableSession.Apply packed", "UIVerifyDesignTableModel packed",
                 "packed.Save3", "UIRerunCloseOutput app, finalPath",
                 "packed.ForceRebuild3", "UIRerunExports packed"]
        positions = [rerun.index(step) for step in order]
        self.assertEqual(positions, sorted(positions))
        for forbidden in ("source.Save3", "MoveFolder", "MoveFile", "ReplaceReferencedDocument", "UIRerunRepairLinks"):
            self.assertNotIn(forbidden, rerun)
        self.assertIn("pack.IncludeDrawings = True", rerun)
        self.assertIn("pack.IncludeSuppressed = True", rerun)
        self.assertIn("includedPaths.Exists(stageFolder & oldNames(i))", rerun)
        self.assertNotIn("pack.AddExternalDocuments(", rerun)
        self.assertIn("CLng(statuses(i)) <> 0", rerun)
        self.assertIn("incomplete save-status list", rerun)
        self.assertIn("fso.FileExists(newNames(i))", rerun)
        self.assertEqual(rerun.count("UIRerunCheckCopiedModel packed, finalFolder"), 2)
        self.assertEqual(rerun.count("app.OpenDoc6"), 3)
        self.assertIn("UIRerunCodeAt", RERUN)
        self.assertIn("Two planter files would both be named", rerun)
        names = helper_body("UIRerunBuildNames", RERUN)
        self.assertIn("serial = serial + 1", names)
        self.assertNotIn("Mid$(oldMaster", names)
        self.assertIn("StrComp(oldMaster", names)
        self.assertIn("UIRerunListModelFiles(stageFolder, oldNames)", rerun)
        self.assertNotIn(".Files", rerun)
        self.assertIn('Left$(fileName, 2) = "~$"', RERUN)
        self.assertIn("UICloseTableEditor(tableSession)", rerun)
        exports = helper_body("UIRunExportsOnly", RERUN)
        for forbidden in ("Save3", "Apply", "MoveFile", "MoveFolder"):
            self.assertNotIn(forbidden, exports)
        detect = helper_body("UIDetectRerunSource", RERUN)
        self.assertIn("On Error GoTo NotOutput", detect)
        self.assertIn("GetDesignTable Is Nothing", detect)

    def test_local_staging_preserves_names_and_limits_cleanup(self):
        source = RERUN.replace("Private Function", "Function")
        stage = helper_body("UIRerunCreateStage", source)
        self.assertIn('Environ("LOCALAPPDATA")', stage.replace("Environ$", "Environ"))
        self.assertIn("If Len(stageRoot) = 0 Then", stage)
        self.assertIn("If Not fso.FolderExists(stageRoot) Then", stage)
        self.assertIn("fso.GetDrive(fso.GetDriveName(stageRoot)).DriveType <> 2", stage)
        self.assertIn('fso.BuildPath(stageRoot, "PlanterRerun-"', stage)
        self.assertNotIn("GetSpecialFolder(2)", stage)
        self.assertIn("fso.GetTempName", stage)
        self.assertIn("fso.CopyFile sourceFolder & names(i), target, False", stage)
        self.assertIn("UIRerunIsInFolder(referencePath, sourceFolder)", stage)
        self.assertIn("app.GetOpenDocumentByName(stagedPath)", stage)
        self.assertIn("app.ReplaceReferencedDocument(stagedPath, referencePath, target)", stage)
        self.assertIn('UITraceDetail "Redirect staged reference"', stage)
        self.assertIn("updatedDependencies = app.GetDocumentDependencies2(stagedPath", stage)
        self.assertIn("StrComp(CStr(updatedDependencies(k + 1)), referencePath, vbTextCompare) = 0", stage)
        self.assertIn("StrComp(CStr(updatedDependencies(k + 1)), target, vbTextCompare) = 0", stage)
        self.assertIn("If Not verified Then", stage)
        self.assertIn("redirectReportedSuccess = app.ReplaceReferencedDocument", stage)
        self.assertNotIn("If Not app.ReplaceReferencedDocument", stage)
        self.assertLess(stage.index("redirectReportedSuccess = app.ReplaceReferencedDocument"),
                        stage.index("updatedDependencies = app.GetDocumentDependencies2"))
        self.assertLess(stage.index("If Not verified Then"),
                        stage.index("If redirectReportedSuccess = False Then"))
        self.assertIn('UITraceDetail "Reference replacement warning"', stage)

        for forbidden in ("MoveFile", "MoveFolder", "GetPackAndGo", "Save3"):
            self.assertNotIn(forbidden, stage)
        cleanup = helper_body("UIRerunCleanStage", source)
        self.assertIn("UIRerunIsInFolder(CStr(path), stageFolder)", cleanup)
        self.assertIn("Kill CStr(path)", cleanup)
        self.assertIn("RmDir Left(stageFolder", cleanup)
        self.assertNotIn("*", cleanup)
        self.assertIn('UITraceDetail "Staging cleanup warning"', cleanup)

    def test_native_booleans_use_zero_comparisons(self):
        for call in (
            "pack.GetDocumentNames(documents)",
            "pack.SetSaveToName(True, finalFolder)",
            "pack.SetDocumentSaveToNames(newNames)",
            "packed.ForceRebuild3(False)",
            "document.Extension.GetWhatsWrong(features, codes, warnings)",
        ):
            self.assertIn(f"If {call} = False Then", RERUN)
            self.assertNotIn(f"If Not {call}", RERUN)

    def test_exporters_accept_parents_without_pack_test(self):
        for name in ("pdfassemblies", "pdfcomponents", "dxfcomponents", "stepcomponents"):
            with self.subTest(exporter=name):
                exporter = (ROOT / "src" / f"{name}.bas").read_text(encoding="utf-8")
                self.assertNotIn('InStr(1, destFolder, "PACK_TEST"', exporter)
                self.assertIn(
                    'If StrComp(Mid$(assemblyPath, InStrRev(assemblyPath, "\\") + 1), '
                    '"planter_assembly.SLDASM", vbTextCompare) = 0 Then',
                    exporter,
                )
                self.assertIn('If assemblyPath = "" Then', exporter)
                self.assertIn('swModel.GetType <> swDocASSEMBLY', exporter)
        root = helper_body("UIOutputRoot")
        self.assertNotIn("PACK_TEST", root)
        self.assertNotIn("packngo1.packTestFolder", FORM)

    def test_assembly_view_checks_use_assigned_positions(self):
        exporter = (ROOT / "src" / "pdfassemblies.bas").read_text(encoding="utf-8")
        body = helper_body("PositionAndScaleAssemblyViews", exporter)
        for view in ("Iso", "Top"):
            for axis in (0, 1):
                self.assertIn(f"verify{view}Pos({axis}) - pos{view}({axis})", body)
            for axis in ("X", "Y"):
                self.assertNotIn(f"target{view}{axis}", body)

    def test_exporter_zero_argument_calls_have_valid_vba_syntax(self):
        exporter = (ROOT / "src" / "pdfassemblies.bas").read_text(encoding="utf-8")
        self.assertIn("Call swDispDim.AutoJogOrdinate()", exporter)
        self.assertNotRegex(exporter, r"(?m)^\s*[A-Za-z_]\w*\.\w+\(\)\s*$")

    def test_form_rerun_controls(self):
        self.assertIn("Private WithEvents exportsOnlyBox As MSForms.CheckBox", FORM)
        self.assertIn('AddCheck("chkExportsOnly"', FORM)
        self.assertIn("If Len(rerunPath) > 0 Then ApplyRerunInputs", FORM)
        apply = helper_body("ApplyRerunInputs", FORM.replace("Private Sub", "Sub"))
        self.assertLess(apply.index("overrideBox_Click"), apply.index("baseWidthBox.Text"))
        run = helper_body("runButton_Click", FORM.replace("Private Sub", "Sub"))
        self.assertLess(run.index("If exportsOnlyBox.Value Then"), run.index("If wallsBox.ListIndex < 0 Then"))
        self.assertIn("exportsOnly = True", run)
        self.assertIn("Choose at least one export for Exports only.", run)
        self.assertIn("Call UIOutputRoot(selectedFolder)", run)
        self.assertIn("fso.GetParentFolderName(fso.GetParentFolderName(rerunPath))", apply)
        rerun = helper_body("UIRunRerun", RERUN)
        self.assertIn("parentRoot = UIOutputRoot(outputFolder)", rerun)
        self.assertLess(rerun.index("parentRoot = UIOutputRoot"), rerun.index("UIRerunSaveOpenWork"))


@unittest.skipUnless(shutil.which("cscript"), "Windows Script Host is required")
class RerunHelpers(unittest.TestCase):
    def test_reference_readback_controls_success_not_api_boolean(self):
        stage = helper_body("UIRerunCreateStage", RERUN.replace("Private Function", "Function"))
        start = stage.index("                    verified = False")
        end = stage.index("                ElseIf Not UIRerunIsInFolder", start)
        verification = stage[start:end]
        script = r'''
Option Explicit
Const RERUN_ERROR = 1001
Dim dependencies, app, warnings
Set app = New FakeApp
Class FakeApp
    Function GetDocumentDependencies2(path, a, b, c)
        GetDocumentDependencies2 = dependencies
    End Function
End Class
Sub UITraceDetail(name, message)
    warnings = warnings + 1
End Sub
Sub Verify(reported)
    Dim verified, updatedDependencies, k, stagedPath, referencePath, target, redirectReportedSuccess
    stagedPath = "stage\drawing.SLDDRW"
    referencePath = "source\part.SLDPRT"
    target = "stage\part.SLDPRT"
    redirectReportedSuccess = reported
''' + verification + r'''
End Sub
Sub Check(reported, expectedFailure, expectedWarnings)
    Dim failure
    warnings = 0
    On Error Resume Next
    Verify reported
    failure = Err.Number
    On Error GoTo 0
    If (failure <> 0) <> expectedFailure Or warnings <> expectedWarnings Then
        WScript.Echo "FAIL: reported=" & reported & "; error=" & failure & "; warnings=" & warnings
        WScript.Quit 1
    End If
End Sub
dependencies = Array("part", "stage\part.SLDPRT")
Check False, False, 1
Check True, False, 0
Check 1, False, 0
If (Not 1) = False Or (1 = False) Then
    WScript.Echo "FAIL: noncanonical native Boolean comparison"
    WScript.Quit 1
End If
dependencies = Array("part", "source\part.SLDPRT")
Check False, True, 0
Check True, True, 0
Check 1, True, 0
dependencies = Array("part", "stage\part.SLDPRT", "old", "source\part.SLDPRT")
Check False, True, 0
Check True, True, 0
dependencies = Empty
Check False, True, 0
Check True, True, 0
dependencies = Array("other", "stage\other.SLDPRT")
Check False, True, 0
Check True, True, 0
WScript.Echo "PASS: reference read-back verification"
'''
        with tempfile.TemporaryDirectory() as folder:
            runner = Path(folder) / "references.vbs"
            runner.write_text(script, encoding="ascii")
            result = subprocess.run(
                ["cscript", "//Nologo", str(runner)], capture_output=True,
                text=True, check=False, timeout=30,
            )
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertIn("PASS:", result.stdout)

    def test_name_round_trip_and_form_inputs(self):
        helpers = "\n\n".join(helper_body(name, RERUN) for name in (
            "UIRerunFormInputs", "UIRerunStripSuffix", "UIRerunAllDigits",
            "UIRerunLengthCode", "UIRerunCodeAt", "UIRerunTemplateFile", "UIRerunIsPlanterFileName",
            "UIRerunNameRole", "UIRerunCurrentBase", "UIRerunRepairEquation",
        ))
        helpers += "\n\n" + "\n\n".join(helper_body(name, PACKNGO) for name in (
            "Step3_BuildNames", "FormatPNDimension", "GetNewBaseName",
        ))
        helpers = re.sub(r"\bLCase\$", "LCase", helpers)
        files = ", ".join(f'"{name}"' for name in TEMPLATE_FILES)
        script = r'''
Option Explicit
Dim wallCount, side, returnType, overallMaterial, thicknessLabel, customSuffix
Dim overallLength, overallWidth, overallHeight, centerPlanterLength, sidePlanterLength
Dim masterName, centerName, centerHardwareName, centerGussetName, mountingGussetName
Dim sideName, sideLeftName, sideRightName, sideCenterName, sideWallName, sharedSideName
Dim sideWallPartName, sideLeftWallPartName, sideRightWallPartName, sideCenterWallPartName
Dim templates, cases, item, suffixes, suffix, packed, roles, result, inputs, bases
SetLocale 1033
Function IIf(condition, a, b)
    If condition Then IIf = a Else IIf = b
End Function
Sub Assert(condition, message)
    If Not condition Then
        WScript.Echo "FAIL: " & message
        WScript.Quit 1
    End If
End Sub
'''
        script += "\n" + helpers + "\n" + rf'''
templates = Array({files})
cases = Array(Array("2", "1", "0", "100", "60", "12", "91.75", "54.5"), _
              Array("3", "0", "3", "120.5", "48", "24.25", "100.25", "40"), _
              Array("4", "0", "1", "60", "60", "18", "60", "60"), _
              Array("1", "0", "2", "36", "12", "8", "36", "0"))
suffixes = Array("", "-C1", "-C12")
For Each item In cases
    For Each suffix In suffixes
        wallCount = item(0) : side = item(1) : returnType = item(2)
        overallLength = item(3) : overallWidth = item(4) : overallHeight = item(5)
        centerPlanterLength = item(6) : sidePlanterLength = item(7)
        overallMaterial = "1" : thicknessLabel = "1"
        customSuffix = ""
        Assert Step3_BuildNames(), "Step3_BuildNames"
        ' Mirror UIApplyActualDimensionNames: side wall parts use a different length than their assembly.
        sideWallPartName = returnType & "-S" & FormatPNDimension(CDbl(sidePlanterLength) - 0.375) & "-" & sideWallName
        sideLeftWallPartName = returnType & "-SL" & FormatPNDimension(CDbl(sidePlanterLength) - 0.375) & "-" & sideWallName
        sideRightWallPartName = returnType & "-SR" & FormatPNDimension(CDbl(sidePlanterLength) - 0.375) & "-" & sideWallName
        sideCenterWallPartName = returnType & "-SC" & FormatPNDimension(CDbl(sidePlanterLength) - 0.375) & "-" & sideWallName
        customSuffix = suffix
        masterName = masterName & suffix
        For Each roles In templates
            packed = GetNewBaseName(roles)
            Assert packed <> "", "No packed name for " & roles
            result = UIRerunTemplateFile(packed, masterName)
            Assert LCase(result) = roles, packed & " mapped to [" & result & "] instead of " & roles
            result = UIRerunTemplateFile(UCase(packed), masterName)
            Assert LCase(result) = roles, "Case-insensitive mapping for " & packed
        Next
    Next
Next
Assert UIRerunTemplateFile("my_bracket.SLDPRT", "20-0-C10000-S6000-1200-11-C1") = "", "Unknown part keeps its name"
Assert UIRerunTemplateFile("0-C9175-1200-11-C1.SLDPRT", "x") = "", "Assembly name on a part file"
Assert UIRerunTemplateFile("0-C9175-1200-11-W.SLDASM", "x") = "", "Part name on an assembly file"
Assert UIRerunTemplateFile("20-0-C10000-S6000-1200-11.SLDPRT", "20-0-C10000-S6000-1200-11") = "", "Master name on a part"
Assert UIRerunTemplateFile("notes.txt", "x") = "", "Non-model file"
Assert Not UIRerunIsPlanterFileName("~$1-1400-22-LW-C1.SLDPRT"), "SolidWorks lock file is not a planter file"
Assert Not UIRerunIsPlanterFileName("notes.txt") And Not UIRerunIsPlanterFileName("README"), "Non-model files"
Assert UIRerunIsPlanterFileName("1-1400-22-LW-C1.SLDPRT") And UIRerunIsPlanterFileName("a.slddrw"), "Model files"
Assert UIRerunNameRole("40-1-C5600-S3600-1400-22-C1") = "planter_assembly", "Master role"
Assert UIRerunNameRole("1-C4200-1400-22-C1") = "planter_center_asm", "Center assembly role"
Assert UIRerunNameRole("1-C4200-1400-22-W-C1") = "planterwallcenter", "Center wall role"
Assert UIRerunNameRole("PLANTER_ASSEMBLY") = "" And UIRerunNameRole("Sketch1") = "", "Non-planter names"
roles = Array("planter_assembly", "planter_center_asm", "planter_side_asm_left", "planterwallcenter")
bases = Array("40-1-C7000-S3600-1400-22-C1", "1-C5600-1400-22-C1", "1-SL3600-1400-22-C1", "1-C5600-1400-22-W-C1")
Assert UIRerunRepairEquation("""OVERALL_LENGTH""= ""OVERALL_LENGTH@1-C4200-1400-22-C1<1>.Assembly""", roles, bases, 4) = _
    """OVERALL_LENGTH""= ""OVERALL_LENGTH@1-C5600-1400-22-C1<1>.Assembly""", "Parent component link"
Assert UIRerunRepairEquation("""Thickness@Sheet-Metal"" = ""OVERALL_THICKNESS@40-1-C5600-S3600-1400-22-C1.Assembly""", roles, bases, 4) = _
    """Thickness@Sheet-Metal"" = ""OVERALL_THICKNESS@40-1-C7000-S3600-1400-22-C1.Assembly""", "Master link"
Assert UIRerunRepairEquation("""1-SL3600-1400-22-C1<3>.Assembly"" = IIF ( ""SIDE"" = 1 , ""suppressed"" , ""unsuppressed"" )", roles, bases, 4) = _
    """1-SL3600-1400-22-C1<3>.Assembly"" = IIF ( ""SIDE"" = 1 , ""suppressed"" , ""unsuppressed"" )", "Current names unchanged"
item = """D1@Sketch2"" = ( ( ""OVERALL_LENGTH"" - ""OVERALL_THICKNESS@PLANTER_ASSEMBLY.Assembly"" ) / ""GUSSETS"" ) + 0.5 / 2"
Assert UIRerunRepairEquation(item, roles, bases, 4) = item, "Template names, sketches and numbers unchanged"
Assert UIRerunRepairEquation("""X"" = ""Y@1-S3600-1400-22-C1.Assembly""", roles, bases, 4) = """X"" = ""Y@1-S3600-1400-22-C1.Assembly""", "Role without current file unchanged"
Assert UIRerunRepairEquation("""X""= ""X@1-c5600-1400-22-c1<1>.Assembly""", roles, bases, 4) = """X""= ""X@1-c5600-1400-22-c1<1>.Assembly""", "Case-only difference unchanged"
item = """OVERALL_LENGTH""= ""OVERALL_LENGTH@1-C5600-1400-22-C1-1.Assembly"""
Assert UIRerunRepairEquation(item, roles, bases, 4) = item, "Current component instance unchanged"
Assert UIRerunRepairEquation("""OVERALL_LENGTH""= ""OVERALL_LENGTH@1-C4200-1400-22-C1-1.Assembly""", roles, bases, 4) = _
    """OVERALL_LENGTH""= ""OVERALL_LENGTH@1-C5600-1400-22-C1-1.Assembly""", "Stale component instance keeps its number"
Assert UIRerunRepairEquation("""X""= ""X@1-C4200-1400-22-2.Assembly""", roles, bases, 4) = _
    """X""= ""X@1-C5600-1400-22-C1-2.Assembly""", "Unsuffixed component instance"
Assert UIRerunRepairEquation("""X""= ""X@1-C4200-1400-22.Assembly""", roles, bases, 4) = _
    """X""= ""X@1-C5600-1400-22-C1.Assembly""", "Unsuffixed document name"
Assert UIRerunTemplateFile("20-0-C10000-S6000-1200-11.SLDASM", "20-0-C10000-S6000-1200-11-C2") = "planter_assembly.sldasm", "Unsuffixed master"
Assert UIRerunStripSuffix("1200-11-CG-C3") = "1200-11-CG", "Strip -C3"
Assert UIRerunStripSuffix("0-C9175-1200-11") = "0-C9175-1200-11", "Keep length code"
Assert UIRerunStripSuffix("1200-11-CG") = "1200-11-CG", "Keep CG"
Assert UIRerunStripSuffix("abc-C") = "abc-C", "Bare C"
Assert UIRerunLengthCode("SL5950") = "SL" And UIRerunLengthCode("S5950") = "S" And UIRerunLengthCode("CG") = "", "Length codes"
inputs = UIRerunFormInputs(Array("2", "2", "3", "100", "60", "12", "2", "0.25", "1", "1", "16.5"))
Assert inputs(0) = "2" And inputs(1) = "Left" And inputs(2) = "Double return", "Rerun walls/side/return"
Assert inputs(3) = 100 And inputs(4) = 60 And inputs(5) = 12, "Rerun dimensions"
Assert inputs(6) = "Borcon Weathering Steel" And inputs(7) = "1/4 inch", "Rerun material"
Assert inputs(8) = True And inputs(9) = True And inputs(10) = 8, "Rerun mounts/override/base width"
inputs = UIRerunFormInputs(Array("3", "1", "0", "100", "60", "12", "1", "0.1875", "0", "0", "0"))
Assert inputs(1) = "N/A" And inputs(2) = "No return" And inputs(6) = "Mild Steel", "Rerun three wall"
Assert inputs(7) = "3/16 inch" And inputs(8) = False And inputs(9) = False And inputs(10) = 0, "Rerun defaults"
inputs = UIRerunFormInputs(Array("2", "1", "1", "100", "60", "12", "1", "0.1875", "0", "0", "0"))
Assert inputs(1) = "Right" And inputs(2) = "Single return 1 inch", "Rerun right side"
WScript.Echo "PASS: rerun helpers"
'''
        with tempfile.TemporaryDirectory() as folder:
            runner = Path(folder) / "rerun.vbs"
            runner.write_text(script, encoding="ascii")
            result = subprocess.run(
                ["cscript", "//Nologo", str(runner)], capture_output=True,
                text=True, check=False, timeout=30,
            )
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertIn("PASS:", result.stdout, result.stdout + result.stderr)


if __name__ == "__main__":
    unittest.main()
