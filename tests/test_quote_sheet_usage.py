"""Verify sheet usage math, form wiring and paginated preview tables."""
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]


class SheetUsageTests(unittest.TestCase):
    def test_form_wiring(self):
        form = (ROOT / "src" / "UserForm_QuoteReport.frm").read_text(encoding="ascii")
        self.assertIn("sheetRows = UIQuoteSheetRows(result)", form)
        self.assertIn("sheetBox.Visible = Not showingDetails", form)
        self.assertIn("sheetTitle.Visible = Not showingDetails", form)
        self.assertIn("UNUSED / POTENTIAL SCRAP", form)
        self.assertIn("includes reusable remnants", form)
        self.assertIn("Me.ScrollBars = fmScrollBarsNone", form)
        self.assertNotIn("Me.ScrollHeight", form)
        self.assertNotIn("Me.ScrollTop", form)
        self.assertIn("Me.Height = 622", form)
        self.assertNotIn("optimizeButton", form)

    @unittest.skipUnless(shutil.which("cscript"), "Windows Script Host required")
    def test_usage_totals_and_pdf_pagination(self):
        def procedure(filename, name):
            source = (ROOT / "src" / filename).read_text(encoding="ascii")
            match = re.search(rf"^(?:Public|Private) (Function|Sub) {name}\b.*?^End \1$",
                              source, re.MULTILINE | re.DOTALL)
            body = re.sub(r" As (?:Long|String|Double|Boolean|Variant|UIQuoteResult)\b", "", match[0])
            body = re.sub(r"^(Public|Private) ", "", body, flags=re.MULTILINE)
            body = re.sub(r"(\w+)\(0 To ([^\n,]+), 0 To ([^\n)]+)\)", r"\1(\2, \3)", body)
            body = re.sub(r"(\w+)\(0 To ([^\n)]+)\)", r"\1(\2)", body)
            body = re.sub(r"^(\s*Next) \w+$", r"\1", body, flags=re.MULTILINE)
            body = body.replace("ByRef offsets()", "ByRef offsets").replace("0#", "0")
            body = body.replace("Right$(", "Right(").replace("Format$(", "PDFFormat(")
            body = body.replace("Mid$(", "Mid(")
            return body

        result_source = (ROOT / "src" / "UIQuoteResult.cls").read_text(encoding="ascii")
        fields = re.sub(r" As (?:Long|String|Double|Boolean|Variant)\b", "",
                        result_source.split("Option Explicit", 1)[1])
        script = "Option Explicit\nSetLocale 1033\nClass UIQuoteResult\n" + fields + "\nEnd Class\n"
        script += r'''
Function UIInchesText(value)
    UIInchesText = CStr(value)
End Function
Function UINestMarginText(result)
    UINestMarginText = CStr(result.EdgeMargin) & " in"
End Function
Function PDFFormat(value, pattern)
    PDFFormat = FormatNumber(value, 2, -1, 0, 0)
End Function
Sub Assert(condition, message)
    If Not condition Then
        WScript.Echo "FAIL: " & message
        WScript.Quit 1
    End If
End Sub
'''
        for name in ("UIQuoteSheetStats", "UIQuoteSheetRows"):
            script += procedure("ui_quote_calculator.bas", name) + "\n"
        for name in ("UIQuoteLayoutPDF", "UIPDFObject", "UIPDFSummary", "UIPDFUsage", "UIPDFOutline", "UIPDFText",
                     "UIPDFStyledText", "UIPDFRightText", "UIPDFRect", "UIPDFHeader", "UIPDFCard", "UIPDFStep", "UIPDFFact"):
            script += procedure("ui_quote_layout_pdf.bas", name) + "\n"
        script += r'''
Function UIPDFParts(result)
    UIPDFParts = UIPDFText(36, 576, 18, "Parts list stub")
End Function
Function UIPDFSheet(result, sheet)
    UIPDFSheet = UIPDFText(36, 576, 18, "Layout sheet " & sheet)
End Function
Dim q, p(24,8), rows, display, i, text, fso, file, code
Set q = New UIQuoteResult
q.NestAvailable = True
q.SheetCount = 25
q.Material = "Mild Steel"
q.Thickness = 0.25
q.PartGap = 0.25
q.EdgeMargin = 0.25
q.EdgeMarginLength = 0.25
q.EdgeMarginWidth = 0.25
q.WallCount = 4
q.Orientation = "N/A"
q.ReturnLabel = "No return"
q.OverallLength = 300
q.OverallWidth = 200
q.Height = 20
q.NestMethod = "MaxRects - mixed gussets"
q.NestTrialCount = 48
q.NestBaselineSheets = 26
For i = 0 To 24
    p(i,0) = i+1
    p(i,6) = 0
    p(i,7) = 144
    p(i,8) = 2
Next
q.SheetPlacements = p
q.TotalArea = 3600
q.UtilizationPercent = 2
q.UnusedPercent = 98
q.AreaMinimumSheets = 1
rows = UIQuoteSheetStats(q)
Assert UBound(rows) = 25 And UBound(rows,2) = 5, "Usage table has per-sheet rows and total"
Assert rows(0,1) = 2 And rows(0,2) = 1 And rows(0,3) = 49, "Pairs count as two physical parts and use true profile area"
Assert rows(25,0) = "TOTAL" And rows(25,1) = 50, "Total physical quantity"
Assert rows(25,2) = 25 And rows(25,3) = 1225, "Total used plus unused equals purchased stock"
Assert rows(25,4) = 2 And rows(25,5) = 98, "Total percentages are weighted, not summed"
display = UIQuoteSheetRows(q)
Assert display(25,4) = "2.00%" And display(25,2) = "25.00", "Formatted display"
text = UIQuoteLayoutPDF(q)
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.CreateTextFile(WScript.Arguments(0),False)
file.Write text
file.Close
q.NestAvailable = False
q.NestStatus = "Sheet estimate withheld: quantities disagree."
display = UIQuoteSheetRows(q)
Assert display(0,0) = "Unavailable" And display(0,3) = "Withheld", "Unavailable is not false zero scrap"
On Error Resume Next
rows = UIQuoteSheetStats(q)
code = Err.Number
Err.Clear
On Error GoTo 0
Assert code <> 0, "Numeric usage unavailable fails explicitly"
q.NestAvailable = True
q.TotalArea = 1
On Error Resume Next
rows = UIQuoteSheetStats(q)
code = Err.Number
Err.Clear
On Error GoTo 0
Assert code <> 0, "Inconsistent aggregate rejected"
WScript.Echo "PASS: usage table totals and unavailable handling"
'''
        with tempfile.TemporaryDirectory(prefix="planter-usage-") as folder:
            runner = Path(folder) / "usage.vbs"
            pdf = Path(folder) / "usage.pdf"
            runner.write_text(script, encoding="ascii")
            run = subprocess.run(["cscript", "//Nologo", str(runner), str(pdf)],
                                 capture_output=True, text=True, timeout=15)
            self.assertEqual(run.returncode, 0, run.stdout + run.stderr)
            self.assertIn("PASS:", run.stdout)
            from pypdf import PdfReader
            reader = PdfReader(pdf, strict=True)
            self.assertEqual(len(reader.pages), 29)
            first = reader.pages[1].extract_text()
            last = reader.pages[2].extract_text()
            self.assertIn("Usage table 1 of 2", first)
            self.assertIn("Usage table 2 of 2", last)
            self.assertNotIn("TOTAL", first)
            self.assertIn("TOTAL", last)
            self.assertIn("1225.00", last)
            self.assertIn("Unused / potential scrap", first)
            self.assertIn("Parts list stub", reader.pages[3].extract_text())
            self.assertIn("Layout sheet 1", reader.pages[4].extract_text())
            self.assertIn("Layout sheet 25", reader.pages[-1].extract_text())
            overview = reader.pages[0].extract_text()
            self.assertIn("Design: 4 walls", overview)
            self.assertIn("MaxRects - mixed gussets", overview)
            self.assertIn("32-trial baseline", overview)
            self.assertIn("26 sheets", overview)
            self.assertIn("Overall: 25 sheets | Utilization 2.00% | Unused / potential scrap 98.00%", last)
            fonts = reader.pages[0]["/Resources"]["/Font"]
            self.assertEqual(fonts["/F2"].get_object()["/BaseFont"], "/Helvetica-Bold")
