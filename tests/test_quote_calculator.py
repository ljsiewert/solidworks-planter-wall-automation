"""Execute standalone VBA calculator with VBScript; no model or Excel required."""
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
CALCULATOR = (ROOT / "src" / "ui_quote_calculator.bas").read_text(encoding="ascii")
MODULE = (ROOT / "src" / "userform_automation.bas").read_text(encoding="ascii")
FORM = (ROOT / "src" / "UserForm_AutomationUI.frm").read_text(encoding="ascii")
RESULT = (ROOT / "src" / "UIQuoteResult.cls").read_text(encoding="ascii")
COMPONENT = (ROOT / "src" / "UIQuoteComponent.cls").read_text(encoding="ascii")


class QuoteTests(unittest.TestCase):
    def test_input_only_routing(self):
        entry = MODULE.split("Public Sub RunWithUI()", 1)[1].split("Public Sub UIDisableExports", 1)[0]
        self.assertLess(entry.index("form.Show vbModal"), entry.index("Step1_Initialize"))
        self.assertLess(entry.index("If quoteOnly Then"), entry.index("Step1_Initialize"))
        self.assertNotIn("Step2_ReadGlobals", entry)
        quote = FORM.split("If quoteOnly Then", 1)[1].split("End If", 1)[0]
        self.assertIn("UICalculateQuote", quote)
        self.assertIn("Exit Sub", quote)
        for flag in ("exportPDFAssemblies", "exportPDFComponents", "exportDXF", "exportSTEP"):
            self.assertIn(flag + " = False", quote)
        for forbidden in ("packngo1", "SldWorks", "UIDesignTableSession", "CreateObject",
                          "OpenDoc", "Save", "Excel", "runLogger"):
            self.assertNotIn(forbidden, CALCULATOR)
        self.assertNotIn("pdfAssembliesBox.Value = True", FORM)
        self.assertIn('256, "0.70"', FORM)
        self.assertIn("pricePerPound = UIQuoteParsePrice(priceBox.Text)", quote)
        self.assertIn("height, pricePerPound)", quote)
        self.assertIn("UIQuoteParsePrice = UIParseInches(text)", CALCULATOR)
        self.assertIn("no currency symbol or commas", CALCULATOR)
        self.assertIn("Optional ByVal pricePerPound As Double = 0.7", CALCULATOR)

    @unittest.skipUnless(shutil.which("cscript"), "Windows Script Host required")
    def test_calculated_values_boundaries_and_report(self):
        def procedure(source, name):
            match = re.search(rf"^(?:Public|Private) (Function|Sub) {name}\b.*?^End \1$",
                              source, re.MULTILINE | re.DOTALL)
            body = re.sub(r" As (?:Long|String|Double|Boolean|Variant|UIQuoteResult|UIQuoteComponent)\b", "", match[0])
            body = body.replace("Private Function", "Function").replace("Public Function", "Function")
            body = body.replace("Private Sub", "Sub")
            body = body.replace("parts(0 To 7)", "parts(7)")
            if name == "UICalculateQuote":
                body = body.replace("Optional ByVal pricePerPound = 0.7", "ByVal pricePerPound")
                body = body.replace("UICalculateQuote", "UIQuoteWithPrice")
            body = re.sub(r"^(\s*Next) \w+$", r"\1", body, flags=re.MULTILINE)
            body = re.sub(r"\bTrim\$", "Trim", body)
            body = re.sub(r"\bMid\$", "Mid", body)
            body = body.replace("Str$(value)", "CStr(value)")
            body = body.replace("Val(cleaned)", "CDbl(cleaned)")
            if name == "UIQuoteParsePrice":
                body = body.replace("On Error GoTo InvalidPrice", "On Error Resume Next\n    Err.Clear")
                body = body.replace("    Exit Function\nInvalidPrice:", "    If Err.Number = 0 Then Exit Function\n    On Error GoTo 0")
            if name == "UIQuoteGussets":
                body = body.replace("    Select Case length\n", "")
                body = body.replace("Case Is <= 24:", "If length <= 24 Then\n")
                for limit in (48, 72, 96):
                    body = body.replace(f"Case Is <= {limit}:", f"ElseIf length <= {limit} Then\n")
                body = body.replace("Case Else:", "Else\n").replace("End Select", "End If")
            return body
        fields = re.sub(r" As (?:Long|String|Double|Boolean|Variant)\b", "",
                        RESULT.split("Option Explicit", 1)[1])
        script = r'''
Option Explicit
SetLocale 1033
Dim q, walls, length, width, thickness, orient, expected, boundary, report, code, i, part, total, ret, h, offset
Dim other, price
Const UI_ERROR = 2100
Sub Assert(condition, message)
    If Not condition Then
        WScript.Echo "FAIL: " & message
        WScript.Quit 1
    End If
End Sub
Function IIf(condition, yes, no)
    If condition Then IIf = yes Else IIf = no
End Function
'''
        script += "\nClass UIQuoteResult\n" + fields + "\nEnd Class\n"
        component_fields = re.sub(r" As (?:String|Double)\b", "", COMPONENT.split("Option Explicit", 1)[1])
        script += "\nClass UIQuoteComponent\n" + component_fields + "\nEnd Class\n"
        for name in ("UICalculateQuote", "UIQuoteFlatHeight", "UIQuotePart",
                     "UIQuoteComponents", "UIQuoteGussets", "UIQuoteReport"):
            script += procedure(CALCULATOR, name) + "\n"
        for name in ("UIValidateSelections", "UIValidateDimensions", "UIInchesText", "UIParseInches"):
            script += procedure(MODULE, name) + "\n"
        script += procedure(CALCULATOR, "UIQuoteParsePrice") + "\n"
        script += r'''
Function UICalculateQuote(walls, orient, ret, material, thickness, length, width, height)
    Set UICalculateQuote = UIQuoteWithPrice(walls, orient, ret, material, thickness, length, width, height, 0.7)
End Function
'''
        script += r'''
Set q = UICalculateQuote(4, "N/A", "Double return", "Mild Steel", "1/4 inch", 300, 450, 48)
Assert q.CenterPatternCount = 3 And q.SidePlantersPerWall = 4, "Screenshot pattern counts"
Assert q.CenterLength = 84 And q.SideLength = 112.875, "Screenshot calculated lengths"
Assert q.CenterGussets = 4 And q.SideGussets = 5, "Screenshot gussets"
Assert q.SideMate = 66 And q.SideMateRight = 234, "Screenshot mate dimensions"
Assert q.CenterWallRuns = 2 And q.SideWallRuns = 2, "Four-wall topology"
report = UIQuoteReport(q)
Assert InStr(report, "Length=112.375; width/height=49.1912") > 0, "Wall dimension labels"
Assert InStr(report, "Length=112.375; width/height=24") > 0, "Base plate dimension labels"
Assert InStr(report, "Length/height=") = 0 And InStr(report, "width/bottom=") = 0, "Old labels removed"
Assert InStr(report, "Derived center planter total: 6") > 0, "Center total"
Assert InStr(report, "Derived side planter total: 8") > 0, "Side total"
Assert InStr(report, "Mounting gussets total: 28") > 0, "Mounting gusset total"
Assert InStr(report, "Center gussets total: 50") > 0, "Internal center gusset total"
Assert q.SideBlankLength = 112.375, "User-confirmed subtraction"
Assert Abs(q.FlatHeight - 49.1912) < 0.000001, "Calibrated double return"
Assert Abs(q.GussetHeight - 45.65) < 0.000001, "Gusset height"
Assert q.GussetBottomWidth = 23.25 And q.GussetTopWidth = 7.75, "Trapezoid widths"
Set part = q.Components(6)
Assert Abs(part.Area - 707.575) < 0.000001 And part.Quantity = 28, "Mounting gusset area and quantity"
Set part = q.Components(7)
Assert part.Quantity = 50, "Center gusset quantity"
Assert Abs(q.TotalArea - 162600.4588) < 0.000001, "Complete reference area"
Assert Abs(q.Density - 0.283599242) < 0.000000001, "7850 kg/m3 converted to lb/in3"
Assert q.PricePerPound = 0.7, "Default price"
Assert Abs(q.TotalWeight - q.TotalArea * q.Thickness * q.Density) < 0.000001, "Profile weight formula"
Assert Abs(q.TotalMaterialCost - q.TotalWeight * 0.7) < 0.000001, "Material cost formula"
Set other = UIQuoteWithPrice(4, "N/A", "Double return", "Borcon Weathering Steel", "1/4 inch", 300, 450, 48, 1.4)
Assert other.Density = q.Density And other.TotalWeight = q.TotalWeight, "Same density for both materials"
Assert Abs(other.TotalMaterialCost - q.TotalMaterialCost * 2) < 0.000001, "User-entered price drives cost"
Set other = UICalculateQuote(4, "N/A", "Single return 1 inch", "Mild Steel", "3/16 inch", 400, 400, 23.75)
Assert Abs(other.TotalArea - 71582.9430666667) < 0.000001, "User example area retained"
Assert Abs(other.TotalWeight - 3806.4128265) < 0.00001, "User example weight"
Assert InStr(report, "NOT CALCULATED") > 0, "No invented cost"
For walls = 1 To 4
    orient = "N/A"
    If walls = 2 Then orient = "Left"
    For Each length In Array(100, 120, 120.5, 121, 240, 300)
        For Each width In Array(100, 120, 120.5, 121, 240, 450)
            For Each thickness In Array("3/16 inch", "1/4 inch")
                Set q = UICalculateQuote(walls, orient, "No return", "Borcon Weathering Steel", thickness, length, width, 10)
                expected = length
                If walls = 2 Then expected = length - 5
                If walls > 2 Then expected = length - 10
                Assert Abs(q.CenterLength * q.CenterSizingSegments - expected) < 0.000001, "Center length conservation"
                Assert q.CenterPatternCount = Int((expected - 1) / 120) + 1, "Assembly count rule"
                Assert q.SidePlantersPerWall = Int((width - 1) / 120) + 1, "Side count rule"
                expected = width / q.SideSizingSegments
                If q.SideSizingSegments > 1 Then expected = expected + q.Thickness + q.Thickness * (q.SideSizingSegments - 2) / q.SideSizingSegments
                Assert Abs(q.SideLength - expected) < 0.000001, "Side thickness adjustment"
                Assert q.CenterWallRuns + q.SideWallRuns = walls, "Wall topology"
                Assert UBound(q.Components) = 7, "Eight physical component groups"
                total = 0
                For i = 0 To 7
                    Set part = q.Components(i)
                    Assert part.Area > 0 And part.Quantity >= 0, "Valid component area and quantity"
                    Assert part.Material = q.Material And part.Thickness = q.Thickness, "Material group preserved"
                    Assert Abs(part.Weight - part.Area * part.Thickness * q.Density) < 0.000001, "Per-part weight"
                    Assert Abs(part.TotalWeight - part.Weight * part.Quantity) < 0.000001, "Component total weight"
                    Assert Abs(part.TotalMaterialCost - part.TotalWeight * q.PricePerPound) < 0.000001, "Component cost"
                    total = total + part.TotalArea
                Next
                If q.HasValidTotals Then
                    Assert Abs(q.TotalArea - total) < 0.000001, "Area sum"
                Else
                    Assert InStr(UIQuoteReport(q), "TOTAL PROFILE AREA") = 0, "No success-shaped area total for mismatch"
                    Assert InStr(UIQuoteReport(q), "TOTAL PROFILE WEIGHT") = 0, "Weight withheld for mismatch"
                    Assert InStr(UIQuoteReport(q), "TOTAL PROFILE MATERIAL COST") = 0, "Cost withheld for mismatch"
                End If
                Set part = q.Components(4)
                Assert part.Quantity = q.SideWallRuns, "Left ends once per run, not per segment"
                Set part = q.Components(5)
                Assert part.Quantity = q.SideWallRuns, "Right ends once per run"
            Next
            For Each thickness In Array("1/4 inch", "3/16 inch")
                For Each h In Array(48, 23.75)
                    For Each ret In Array("No return", "Single return 1 inch", "Single return 2 inches", "Double return")
                        Set q = UICalculateQuote(1, "N/A", ret, "Mild Steel", thickness, 100, 100, h)
                        offset = 0
                        If ret = "Single return 1 inch" Or ret = "Single return 2 inches" Then
                            offset = IIf(thickness = "1/4 inch", 0.5956, 0.6798)
                        End If
                        If ret = "Single return 2 inches" Then offset = offset + 1
                        If ret = "Double return" Then offset = IIf(thickness = "1/4 inch", 1.1912, 1.3595)
                        Assert Abs(q.FlatHeight - (h + offset)) < 0.000001, "All measured and predicted flat heights"
                        Set part = q.Components(2)
                        Assert part.Quantity = 0, "One-wall side wall absent"
                        Set part = q.Components(6)
                        Assert part.Quantity = 2 * q.CenterPatternCount, "One-wall mounting gussets"
                    Next
                Next
            Next
        Next
    Next
Next
For Each boundary In Array(24, 48, 72, 96)
    Set q = UICalculateQuote(1, "N/A", "No return", "Mild Steel", "3/16 inch", boundary, 100, 6)
    Assert q.CenterGussets = boundary / 24, "Gusset threshold inclusive"
    Set q = UICalculateQuote(1, "N/A", "No return", "Mild Steel", "3/16 inch", boundary + 0.001, 100, 6)
    Assert q.CenterGussets = boundary / 24 + 1, "Gussets above threshold"
Next
Set q = UICalculateQuote(1, "N/A", "No return", "Mild Steel", "3/16 inch", 120.5, 120, 10)
Assert q.CenterSizingSegments = 2 And q.CenterPatternCount = 1, "Preserve source boundary discrepancy"
report = UIQuoteReport(q)
Assert InStr(report, "REVIEW REQUIRED") > 0 And InStr(report, "Derived center planter total:") = 0, "Withhold unreliable totals"
Set q = UICalculateQuote(1, "N/A", "No return", "Mild Steel", "3/16 inch", 100, 120.5, 10)
Assert q.SideWallRuns = 0 And Len(q.Warnings) = 0, "Unused side discrepancy does not invalidate one wall"
For Each orient In Array("Left", "Right")
    Set q = UICalculateQuote(2, orient, "Single return 1 inch", "Mild Steel", "1/4 inch", 200, 100, 10)
    Assert q.SideWallRuns = 1 And q.CenterWallRuns = 1, "Two-wall orientation"
Next
On Error Resume Next
Set q = UICalculateQuote(5, "N/A", "No return", "Mild Steel", "1/4 inch", 100, 100, 10)
code = Err.Number
On Error GoTo 0
Assert code <> 0, "Invalid selections rejected"
For Each price In Array(0, -0.7)
    On Error Resume Next
    Set q = UIQuoteWithPrice(1, "N/A", "No return", "Mild Steel", "1/4 inch", 100, 100, 10, price)
    code = Err.Number
    Err.Clear
    On Error GoTo 0
    Assert code <> 0, "Nonpositive price rejected"
Next
For Each price In Array("0.70", ".7", "1.25", " 0.70 ")
    Assert UIQuoteParsePrice(price) = CDbl(Trim(price)), "Valid decimal pricing input"
Next
For Each price In Array("", "0", "-1", "$0.70", "0,70", "1e2", "1/2", "0.7.0")
    On Error Resume Next
    expected = UIQuoteParsePrice(price)
    code = Err.Number
    Err.Clear
    On Error GoTo 0
    Assert code <> 0, "Invalid price string rejected"
Next
WScript.Echo "PASS: screenshot, 288 scenarios, thresholds, warnings, totals and invalid inputs"
'''
        with tempfile.TemporaryDirectory(prefix="planter-quote-") as folder:
            runner = Path(folder) / "quote.vbs"
            runner.write_text(script, encoding="ascii")
            result = subprocess.run(["cscript", "//Nologo", str(runner)], capture_output=True,
                                    text=True, timeout=30)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertIn("PASS:", result.stdout, result.stdout + result.stderr)
