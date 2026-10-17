"""Execute standalone VBA calculator with VBScript; no model or Excel required."""
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import unittest
import math

ROOT = Path(__file__).resolve().parents[1]
CALCULATOR = (ROOT / "src" / "ui_quote_calculator.bas").read_text(encoding="ascii")
MODULE = (ROOT / "src" / "userform_automation.bas").read_text(encoding="ascii")
FORM = (ROOT / "src" / "UserForm_AutomationUI.frm").read_text(encoding="ascii")
RESULT = (ROOT / "src" / "UIQuoteResult.cls").read_text(encoding="ascii")
COMPONENT = (ROOT / "src" / "UIQuoteComponent.cls").read_text(encoding="ascii")
REPORT_FORM = (ROOT / "src" / "UserForm_QuoteReport.frm").read_text(encoding="ascii")
NESTING = (ROOT / "src" / "ui_quote_nesting.bas").read_text(encoding="ascii")
PDF = (ROOT / "src" / "ui_quote_layout_pdf.bas").read_text(encoding="ascii")

def polygons_overlap(first, second):
    """Separating-axis test for the convex profiles drawn in the preview."""
    for polygon in (first, second):
        for i, (x, y) in enumerate(polygon):
            xx, yy = polygon[(i + 1) % len(polygon)]
            axis = (y - yy, xx - x)
            a = [px * axis[0] + py * axis[1] for px, py in first]
            b = [px * axis[0] + py * axis[1] for px, py in second]
            if max(a) <= min(b) + 1e-9 or max(b) <= min(a) + 1e-9:
                return False
    return True


def polygon_clearance(first, second):
    def point_segment(point, start, end):
        dx, dy = end[0] - start[0], end[1] - start[1]
        t = max(0, min(1, ((point[0] - start[0]) * dx +
                          (point[1] - start[1]) * dy) / (dx * dx + dy * dy)))
        return math.hypot(point[0] - start[0] - t * dx,
                          point[1] - start[1] - t * dy)
    return min(point_segment(point, polygon[i], polygon[(i + 1) % len(polygon)])
               for points, polygon in ((first, second), (second, first))
               for point in points for i in range(len(polygon)))


class QuoteTests(unittest.TestCase):
    def test_summary_form_and_details_wiring(self):
        self.assertNotRegex(REPORT_FORM, r"\bheading[ \t]+\.")
        self.assertIn("heading.Top = 188", REPORT_FORM)
        set_result = REPORT_FORM.split("Public Sub SetResult(", 1)[1].split("End Sub", 1)[0]
        self.assertIn("FillTable(componentsBox, componentRows", set_result)
        self.assertIn("FillTable(gussetsBox, gussetRows", set_result)
        self.assertIn("FillTable sheetBox, sheetRows", set_result)
        fill = REPORT_FORM.split("Private Function FillTable(", 1)[1].split("End Function", 1)[0]
        self.assertIn("table.Height = contentHeight + 4", fill, "Tables are exactly as tall as their rows")
        self.assertIn("contentHeight = (UBound(rows, 1) + 1) * ROW_HEIGHT", fill)
        self.assertIn("table.ScrollBars = fmScrollBarsVertical", fill, "Long tables scroll internally")
        self.assertNotIn("ListShowsAllRows", REPORT_FORM)
        self.assertNotIn('"Forms.ListBox.1", "lstComponents"', REPORT_FORM)
        self.assertIn("report.SetResult result", CALCULATOR)
        self.assertIn("cards = UIQuoteSummaryCards(result)", REPORT_FORM)
        self.assertIn("tileValues(i).Caption = cards(i + 2, 1)", REPORT_FORM)
        self.assertIn("componentRows = UIQuoteComponentRows(result)", REPORT_FORM)
        self.assertIn("gussetRows = UIQuoteGussetRows(result)", REPORT_FORM)
        self.assertIn("totalsBox.List = UIQuoteTotalRows(result)", REPORT_FORM)
        self.assertIn("gussetsBox.Visible = Not showingDetails", REPORT_FORM)
        self.assertIn("totalsBox.Visible = Not showingDetails", REPORT_FORM)
        self.assertIn("Private WithEvents pdfButton", REPORT_FORM)
        self.assertIn("pdfButton.Enabled = result.NestAvailable", REPORT_FORM)
        self.assertIn("UIPreviewQuoteLayoutPDF quoteResult", REPORT_FORM)
        self.assertIn("CreateTextFile(path, False, False)", PDF)
        self.assertIn('shell.ShellExecute path, "", "", "open", 1', PDF)
        self.assertIn("fso.GetSpecialFolder(2)", PDF)
        preview_handler = REPORT_FORM.split("Private Sub pdfButton_Click()", 1)[1].split("End Sub", 1)[0]
        self.assertNotIn("BrowseForFolder", preview_handler)
        self.assertIn("reportBox.Text = UIQuoteReport(result)", REPORT_FORM)
        self.assertIn("Private WithEvents detailsButton", REPORT_FORM)
        self.assertIn("Private Sub detailsButton_Click()", REPORT_FORM)
        self.assertIn("reportBox.Visible = showingDetails", REPORT_FORM)
        self.assertIn("heading.Visible = Not showingDetails", REPORT_FORM)
        self.assertIn("item.Visible = Not showingDetails", REPORT_FORM)
        self.assertNotIn("summaryBox", REPORT_FORM)
        self.assertIn("componentsBox.Visible = Not showingDetails", REPORT_FORM)
        for warning in ("not a final manufacturing quote", "BORCON-specific density is unverified",
                        "Holes/reliefs not deducted", "other heights are predictions"):
            self.assertIn(warning, REPORT_FORM)

    def test_48_heuristic_trials_baseline32_and_flexible_gussets(self):
        self.assertIn("For mode = 0 To 27", NESTING)
        self.assertIn("If mode < 4 Or packMode >= 2 Or sortOrder >= 2 Then", NESTING)
        trials = []
        for mode in range(28):
            if mode < 4:
                sort_order = mode % 2
                split_mode = mode // 2
            else:
                sort_order = (mode - 4) % 4
                split_mode = (mode - 4) // 4
            if mode < 4 or split_mode >= 2 or sort_order >= 2:
                trials.append((sort_order, split_mode))
        self.assertEqual(trials[:4], [(0, 0), (1, 0), (0, 1), (1, 1)])
        self.assertEqual(len(trials), 24)
        self.assertEqual(set(trials), {(order, split) for order in range(4) for split in range(6)})
        # Baseline MaxRects: 4 sort orders x 2 rect heuristics = 8 trials
        # Total baseline = 24 + 8 = 32 trials
        # Flexible gussets: 2 policies (single=1, mixed=2) x 4 sort orders x 2 rect heuristics = 16 trials
        # Total evaluated candidate layouts = 32 + 16 = 48 trials
        self.assertIn("result.NestTrialCount = 48", NESTING)
        self.assertIn("singleItems = UINestBuildItems(result, 1", NESTING)
        self.assertIn("mixedItems = UINestBuildItems(result, 2", NESTING)

    def test_input_only_routing(self):
        entry = MODULE.split("Public Sub RunWithUI()", 1)[1].split("Public Sub UIDisableExports", 1)[0]
        self.assertLess(entry.index("form.Show vbModal"), entry.index("Step1_Initialize"))
        self.assertLess(entry.index("If quoteOnly Then"), entry.index("Step1_Initialize"))
        self.assertNotIn("Step2_ReadGlobals", entry)
        quote = FORM.split("Private Function ReadQuoteDesign()", 1)[1].split("End Function", 1)[0]
        self.assertIn("UICalculateQuote", quote)
        self.assertIn("Exit Function", quote)
        run = FORM.split("Private Sub runButton_Click()", 1)[1].split("End Sub", 1)[0]
        for flag in ("exportPDFAssemblies", "exportPDFComponents", "exportDXF", "exportSTEP"):
            self.assertIn(flag + " = False", run)
        for forbidden in ("packngo1", "SldWorks", "UIDesignTableSession", "CreateObject",
                          "OpenDoc", "Save", "Excel", "runLogger"):
            self.assertNotIn(forbidden, CALCULATOR)
        self.assertNotIn("pdfAssembliesBox.Value = True", FORM)
        self.assertIn('AddPropertyText("txtPricePerPound", "Quote material ($/lb)", 40, "0.90")', FORM)
        self.assertIn('priceBox.Text = "0.90"', FORM)
        self.assertIn("typical range 0-0.25 in", FORM)
        self.assertIn("UIQuotePricing result, CBool(paintedToggle.Value)", quote)
        self.assertIn('"Forms.ToggleButton.1", "tglPainted"', FORM)
        self.assertIn('AddButton("cmdProperties", "Properties..."', FORM)
        for box in ("txtCutRate", "txtBendRate", "txtWeldRate", "txtPaintRate", "txtQuoteMargin",
                    "txtPartGap", "txtEdgeMargin"):
            self.assertIn(f'AddPropertyText("{box}"', FORM)
        self.assertIn("Private WithEvents marginSpin As MSForms.SpinButton", REPORT_FORM)
        self.assertIn("UIQuoteSetMargin quoteResult, margin", REPORT_FORM)
        self.assertIn("pricing = UIQuotePricingCards(quoteResult)", REPORT_FORM)
        self.assertIn("pricePerPound = UIQuoteParsePrice(priceBox.Text)", quote)
        self.assertIn("height, pricePerPound, partGap, edgeMargin)", quote)
        self.assertIn("UIQuoteParsePrice = UIParseInches(text)", CALCULATOR)
        self.assertIn("no currency symbol or commas", CALCULATOR)
        self.assertIn("Optional ByVal pricePerPound As Double = 0.9", CALCULATOR)

    def test_field_assembly_wiring(self):
        field_form = (ROOT / "src" / "UserForm_FieldQuote.frm").read_text(encoding="ascii")
        self.assertIn('Me.Controls.Add("Forms.TextBox.1", "txtQuantity", True)', FORM)
        self.assertNotIn('AddText("txtQuantity"', FORM)
        self.assertIn("quantityBox.Enabled = quoteBox.Value", FORM)
        self.assertIn("fieldDesigns.Add result", FORM)
        self.assertIn("fieldDesigns.Remove planterList.ListIndex + 1", FORM)
        self.assertIn("Set quoteDesigns = fieldDesigns", FORM)
        self.assertIn("UIQuoteSetQuantity result, quantity", FORM)
        self.assertIn("UIShowFieldQuote quoteDesigns", MODULE)
        self.assertEqual(MODULE.count("Set quoteDesigns = Nothing"), 2)
        run = FORM.split("Private Sub runButton_Click()", 1)[1].split("End Sub", 1)[0]
        self.assertLess(run.index("Set quoteDesigns = fieldDesigns"), run.index("UIValidateSelections"))
        self.assertIn('If fieldDesigns.Count = 0 Then', run)
        self.assertNotIn("UICalculateQuote", run)
        self.assertIn("UIFieldQuoteTotals(designs)", field_form)
        self.assertIn("UIShowQuote result, False\n    RefreshTotals", field_form)
        self.assertIn("UIFieldSetMargin designs, margin", field_form)
        self.assertIn("Private WithEvents marginSpin As MSForms.SpinButton", field_form)
        self.assertNotIn("marginButton", field_form)
        self.assertIn("Set previewResult = designs(designList.ListIndex + 1)", field_form)
        self.assertIn("FIELD TOTALS WITHHELD", field_form)
        self.assertIn('widths = Array(360, 55, 65, 120, 130)', field_form)
        self.assertIn("label.Left = x + 3", field_form)
        self.assertIn("x = x + widths(i)", field_form)
        self.assertIn(".ColumnWidths = columnWidths", field_form)
        self.assertNotIn('"lblFieldColumns"', field_form)

    def test_pdf_quantity_line_clears_summary_cards(self):
        quantity_y = float(re.search(
            r'UIPDFStyledText\(36, (\d+), 9, result.EstimateLabel', PDF)[1])
        card = re.search(r"Private Function UIPDFCard\b.*?End Function", PDF, re.DOTALL)[0]
        bottom, height = map(float, re.search(r"UIPDFRect\(x, (\d+), 171, (\d+),", card).groups())
        caption_y = float(re.search(r"UIPDFStyledText\(x \+ 14, (\d+), 8.5, caption", card)[1])
        self.assertGreaterEqual(quantity_y - (bottom + height), 12)
        self.assertGreaterEqual(bottom + height - caption_y, 12)

    @unittest.skipUnless(shutil.which("cscript"), "Windows Script Host required")
    def test_calculated_values_boundaries_and_report(self):
        def procedure(source, name):
            match = re.search(rf"^(?:Public|Private) (Function|Sub) {name}\b.*?^End \1$",
                              source, re.MULTILINE | re.DOTALL)
            body = re.sub(r" As (?:Long|String|Double|Boolean|Variant|Object|UIQuoteResult|UIQuoteComponent)\b", "", match[0])
            body = body.replace("Private Function", "Function").replace("Public Function", "Function")
            body = body.replace("Private Sub", "Sub")
            body = body.replace("Public Sub", "Sub")
            body = body.replace("parts(0 To 7)", "parts(7)")
            body = body.replace("rows(0 To count - 1, 0 To weightColumn + 1)", "rows(count - 1, weightColumn + 1)")
            body = body.replace("rows(0 To 0, 0 To 2)", "rows(0, 2)")
            body = re.sub(r"(\w+)\(0 To ([^\n,]+), 0 To ([^\n)]+)\)", r"\1(\2, \3)", body)
            body = re.sub(r"(\w+)\(0 To ([^\n)]+)\)", r"\1(\2)", body)
            body = body.replace("usedSheets(1 To sheets)", "usedSheets(sheets)")
            if name == "UICalculateQuote":
                body = body.replace("Optional ByVal pricePerPound = 0.9", "ByVal pricePerPound")
                body = body.replace("Optional ByVal partGap = 0.25", "ByVal partGap")
                body = body.replace("Optional ByVal edgeMargin = 0.25", "ByVal edgeMargin")
                body = body.replace("UICalculateQuote", "UIQuoteWithPrice")
            body = re.sub(r"^(\s*Next) \w+$", r"\1", body, flags=re.MULTILINE)
            body = re.sub(r"\bTrim\$", "Trim", body)
            body = re.sub(r"\bMid\$", "Mid", body)
            body = body.replace("Right$(", "Right(").replace("Format$(", "FormatNumberPDF(")
            body = body.replace("Left$(", "Left(")
            body = body.replace("ByRef offsets()", "ByRef offsets").replace("ByRef gussets()", "ByRef gussets").replace("0#", "0")
            body = body.replace("Str$(value)", "CStr(value)")
            body = body.replace("Val(cleaned)", "CDbl(cleaned)")
            if name == "UISaveQuoteLayoutPDF":
                body = body.split("Failed:", 1)[0].replace("On Error GoTo Failed", "") + "End Sub"
            if name == "UIPreviewQuoteLayoutPDF":
                body = body.replace('CreateObject("Shell.Application")', "New LayoutShell")
            if name == "UIPDFSheet":
                body = body.replace("Dim quantities(7)", "Dim quantities(7)\n    For i = 0 To 7\n        quantities(i) = 0\n    Next")
            if name in ("UIQuoteParsePrice", "UIQuoteParseRate"):
                label = "InvalidPrice" if name == "UIQuoteParsePrice" else "InvalidRate"
                body = body.replace(f"On Error GoTo {label}", "On Error Resume Next\n    Err.Clear")
                body = body.replace(f"    Exit Function\n{label}:", "    If Err.Number = 0 Then Exit Function\n    On Error GoTo 0")
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
Dim other, price, summary, rows, gussetRows, totalRows, originalArea, originalWeight, originalCost
Dim p, r, j, dx, dy, sheetArea, physicalCount, gap, margin, pairedCount, customItems(0, 4), customSheets
Dim heuristicItems(7, 4), trialItems, trialWidths, trialHeights, legacyBest, extraBest, trialSheets, sortOrder, splitMode
Dim maxItems(9, 4), maxBest, seenMax, stressItems(95, 4)
Dim pdfText, pdfFile, pdfFSO, previewPath, relaxLength, relaxWidth, relaxNames
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
Function FormatNumberPDF(value, pattern)
    FormatNumberPDF = FormatNumber(value, 2, -1, 0, 0)
End Function
Class LayoutShell
    Sub ShellExecute(path, args, directory, action, show)
        Assert action = "open", "Preview uses default viewer"
        Assert Right(path, 4) = ".pdf", "Preview launches PDF"
        previewPath = path
    End Sub
End Class
'''
        script += "\nClass UIQuoteResult\n" + fields + "\nEnd Class\n"
        component_fields = re.sub(r" As (?:String|Double)\b", "", COMPONENT.split("Option Explicit", 1)[1])
        script += "\nClass UIQuoteComponent\n" + component_fields + "\nEnd Class\n"
        for name in ("UICalculateQuote", "UIQuoteFlatHeight", "UIQuotePart",
                     "UIQuoteComponents", "UIQuoteGussets", "UIQuoteReport",
                     "UIQuoteSummaryCards", "UIQuoteComponentRows", "UIQuoteGussetRows",
                     "UIQuoteTableRows", "UIQuoteTotalRows", "UIQuoteSheetStats", "UIQuoteSheetRows",
                     "UIQuoteDefaultRate", "UIQuotePricing", "UIQuoteSetMargin", "UIQuoteParseRate",
                     "UIQuotePricingCards", "UIQuotePricingReport"):
            script += procedure(CALCULATOR, name) + "\n"
        for name in ("UIValidateSelections", "UIValidateDimensions", "UIInchesText", "UIParseInches"):
            script += procedure(MODULE, name) + "\n"
        for name in ("UIQuoteParseQuantity", "UIQuoteSetQuantity"):
            script += procedure(CALCULATOR, name) + "\n"
        field = (ROOT / "src" / "ui_quote_field.bas").read_text(encoding="ascii")
        script += procedure(field, "UIFieldQuoteTotals").replace(
            " As Collection", "").replace("For Each result In designs", "For Each result In designs.Items") + "\n"
        script += procedure(CALCULATOR, "UIQuoteParsePrice") + "\n"
        for name in ("UINestQuote", "UINestSort", "UINestMax", "UINestPack", "UINestMaxRects",
                     "UINestFreeAdd", "UINestFreePrune", "UINestReport", "UINestBuildItems",
                     "UINestFreeSubtract", "UINestEliminateSheet", "UINestRelaxMargin", "UINestMarginText"):
            script += procedure(NESTING, name) + "\n"
        for name in ("UIQuoteLayoutPDF", "UIPDFObject", "UIPDFSummary", "UIPDFUsage", "UIPDFSheet", "UIPDFComponentIndex", "UIPDFOutline",
                     "UIPDFText", "UIPDFStyledText", "UIPDFRightText", "UIPDFRect", "UIPDFHeader", "UIPDFCard",
                     "UIPDFStep", "UIPDFFact", "UIPDFParts", "UIPDFPartCounts", "UIPDFPartLabel", "UISaveQuoteLayoutPDF", "UIPreviewQuoteLayoutPDF"):
            script += procedure(PDF, name) + "\n"
        script += r'''
Dim copies, fieldDesigns, fieldTotals, baseSheets, basePaint, baseLabor, baseArea, baseQuantity
Set copies = UICalculateQuote(1, "N/A", "No return", "Mild Steel", "1/4 inch", 36, 24, 12)
UIQuotePricing copies, True, 75, 5, 10, 2.5, 50
baseArea = copies.TotalArea
basePaint = copies.PaintCost
baseQuantity = copies.WeldedParts
baseSheets = copies.SheetCount
UIQuoteSetQuantity copies, 3
Assert copies.Quantity = 3 And Abs(copies.TotalArea - baseArea * 3) < 0.000001, "Quantity scales profile totals"
Assert copies.WeldedParts = baseQuantity * 3 And copies.WeldCost = 10 * baseQuantity * 3, "Quantity scales physical parts and welding"
Assert Abs(copies.PaintCost - basePaint * 3) < 0.000001, "Quantity scales both-sided paint"
Assert copies.SheetCount <= baseSheets * 3 And copies.CutCost = 75 * copies.SheetCount, "All copies re-nested and cutting priced on actual sheets"
CheckNest copies
Assert InStr(UIQuoteReport(copies), "quantity: 3") > 0, "Technical report identifies quantity"
Assert InStr(UIQuoteLayoutPDF(copies), "Quantity: 3") > 0, "PDF identifies quantity"
Assert InStr(UIQuoteReport(copies), "Derived center planter total: " & UIInchesText(3)) > 0, "Derived counts include copies"
UIQuoteSetQuantity copies, 2
Assert copies.Quantity = 2 And Abs(copies.TotalArea - baseArea * 2) < 0.000001, "Quantity changes are absolute, not compounded"
CheckNest copies
UIQuoteSetQuantity copies, 1
Assert copies.WeldedParts = baseQuantity And Abs(copies.TotalArea - baseArea) < 0.000001, "Restoring one copy preserves original totals"
On Error Resume Next
UIQuoteSetQuantity copies, 0
code = Err.Number
Err.Clear
On Error GoTo 0
Assert code <> 0 And copies.Quantity = 1, "Invalid quantity leaves result unchanged"
UIQuoteSetQuantity copies, 2000
Assert Not copies.HasPricing And Not copies.NestAvailable And InStr(copies.NestStatus, "2000 physical components") > 0, "Physical component cap applies after multiplying copies"
UIQuoteSetQuantity copies, 1
Assert copies.HasPricing, "Returning to a supported quantity restores pricing"
Assert UIQuoteParseQuantity(" 2 ") = 2 And UIQuoteParseQuantity("2000") = 2000, "Whole-number quantity boundaries"
For Each price In Array("", "0", "-1", "1.5", "1e2", "2,000", "2001", "99999", "two")
    On Error Resume Next
    i = UIQuoteParseQuantity(CStr(price))
    code = Err.Number
    Err.Clear
    On Error GoTo 0
    Assert code <> 0, "Invalid quantity rejected: " & price
Next
Set fieldDesigns = CreateObject("Scripting.Dictionary")
fieldDesigns.Add "A", copies
Set other = UICalculateQuote(1, "N/A", "No return", "Borcon Weathering Steel", "3/16 inch", 48, 24, 12)
UIQuoteSetQuantity other, 2
fieldDesigns.Add "B", other
fieldTotals = UIFieldQuoteTotals(fieldDesigns)
Assert fieldTotals(8) And fieldTotals(0) = 3, "Field totals include all planter copies"
Assert fieldTotals(1) = copies.SheetCount + other.SheetCount, "Separate designs keep separate sheet counts"
Assert Abs(fieldTotals(5) - copies.TotalCost - other.TotalCost) < 0.000001, "Field total cost sums designs"
Assert Abs(fieldTotals(7) - copies.SellPrice - other.SellPrice) < 0.000001, "Field sell price sums designs"
UIQuoteSetMargin copies, 40
fieldTotals = UIFieldQuoteTotals(fieldDesigns)
Assert Abs(fieldTotals(7) - copies.SellPrice - other.SellPrice) < 0.000001, "Field totals reflect design margin edits"
other.HasPricing = False
fieldTotals = UIFieldQuoteTotals(fieldDesigns)
Assert Not fieldTotals(8), "Unpriced design withholds complete field totals"
fieldDesigns.RemoveAll
On Error Resume Next
fieldTotals = UIFieldQuoteTotals(fieldDesigns)
code = Err.Number
Err.Clear
On Error GoTo 0
Assert code <> 0, "Empty field quote rejected"
Function UICalculateQuote(walls, orient, ret, material, thickness, length, width, height)
    Set UICalculateQuote = UIQuoteWithPrice(walls, orient, ret, material, thickness, length, width, height, 0.9, 0.25, 0.25)
End Function
Sub CheckNest(q)
    Dim p, i, j, dx, dy, area, parts, pairs, gap, margin, occurrences, labels
    Assert q.NestAvailable And q.SheetCount > 0, "Nesting available"
    p = q.SheetPlacements
    gap = q.PartGap
    margin = q.EdgeMargin
    area = 0
    parts = 0
    pairs = 0
    For i = 0 To UBound(p)
        Assert p(i, 0) >= 1 And p(i, 0) <= q.SheetCount, "Placement sheet"
        Assert p(i, 2) >= q.EdgeMarginLength - 0.000001 And p(i, 3) >= q.EdgeMarginWidth - 0.000001, "Near edge margin"
        Assert p(i, 2) + p(i, 4) <= 120 - q.EdgeMarginLength + 0.000001, "Far length margin"
        Assert p(i, 3) + p(i, 5) <= 60 - q.EdgeMarginWidth + 0.000001, "Far width margin"
        Assert p(i, 6) = 0 Or p(i, 6) = 90, "Allowed bundle rotation"
        area = area + p(i, 7)
        parts = parts + p(i, 8)
        If p(i, 8) = 2 Then
            pairs = pairs + 1
            dx = gap * Sqr(1 + ((q.GussetBottomWidth - q.GussetTopWidth) / q.GussetHeight)^2)
            Assert Abs(dx * q.GussetHeight / Sqr(q.GussetHeight^2 + (q.GussetBottomWidth-q.GussetTopWidth)^2) - gap) < 0.000001, "Gusset pair perpendicular gap"
        End If
        For j = i + 1 To UBound(p)
            If p(i, 0) = p(j, 0) Then
                Assert p(i, 2) + p(i, 4) + gap <= p(j, 2) + 0.000001 Or _
                       p(j, 2) + p(j, 4) + gap <= p(i, 2) + 0.000001 Or _
                       p(i, 3) + p(i, 5) + gap <= p(j, 3) + 0.000001 Or _
                       p(j, 3) + p(j, 5) + gap <= p(i, 3) + 0.000001, "No overlap; minimum bundle gap"
            End If
        Next
    Next
    Assert pairs > 0, "Shape-aware pairs placed"
    dx = 0
    For i = 0 To UBound(q.Components)
        dx = dx + q.Components(i).Quantity
    Next
    Assert parts = dx, "All physical components placed exactly once"
    For i = 0 To UBound(q.Components)
        occurrences = 0
        For j = 0 To UBound(p)
            labels = Split(Replace(p(j, 1), " (opposed pair)", ""), " + ")
            For Each dx In labels
                If dx = q.Components(i).Name Then occurrences = occurrences + 1
            Next
        Next
        Assert occurrences = q.Components(i).Quantity, "Each component type quantity conserved"
    Next
    Assert Abs(area - q.TotalArea) < 0.000001, "Profile areas conserved, not envelope areas"
    Assert q.SheetCount >= q.AreaMinimumSheets, "Sheet count above area lower bound"
    Assert Abs(q.UtilizationPercent + q.UnusedPercent - 100) < 0.000001, "Scrap complement"
    Assert Abs(q.UtilizationPercent - 100 * area / (7200 * q.SheetCount)) < 0.000001, "Full-sheet utilization basis"
    Assert q.NestTrialCount = 48, "48 candidate layouts evaluated"
    Assert Len(q.NestMethod) > 0, "NestMethod populated"
    Assert q.NestBaselineSheets >= q.SheetCount, "Sheet count no worse than baseline"
    CheckNoRegression q
End Sub
Sub CheckNoRegression(q)
    Dim items, totalGussets, i, j, k, part, pairFits, pairWidth, dx, gussets(), count, baseline, sortOrder, splitMode, sheets, trial
    Assert q.SheetCount <= q.NestBaselineSheets, "Nesting is no worse than the prior 32 baseline layouts"
    If q.SheetCount <= q.AreaMinimumSheets And q.NestBaselineSheets = q.SheetCount Then
        Exit Sub
    End If
    totalGussets = q.Components(6).Quantity + q.Components(7).Quantity
    pairWidth = q.GussetBottomWidth + q.GussetTopWidth + q.PartGap * Sqr(1 + ((q.GussetBottomWidth - q.GussetTopWidth) / q.GussetHeight)^2)
    pairFits = (pairWidth <= 120 - 2 * q.EdgeMarginLength And q.GussetHeight <= 60 - 2 * q.EdgeMarginWidth) Or _
               (q.GussetHeight <= 120 - 2 * q.EdgeMarginLength And pairWidth <= 60 - 2 * q.EdgeMarginWidth)
    count = 0
    For i = 0 To 5
        count = count + q.Components(i).Quantity
    Next
    If pairFits Then
        count = count + Int(totalGussets / 2) + (totalGussets Mod 2)
    Else
        count = count + totalGussets
    End If
    ReDim items(count - 1, 4)
    k = 0
    For i = 0 To 5
        Set part = q.Components(i)
        For j = 1 To part.Quantity
            items(k, 0) = part.Name
            items(k, 1) = part.Length
            items(k, 2) = part.Width
            items(k, 3) = part.Area
            items(k, 4) = 1
            k = k + 1
        Next
    Next
    ReDim gussets(totalGussets)
    j = 0
    For i = 1 To q.Components(6).Quantity
        gussets(j) = q.Components(6).Name
        j = j + 1
    Next
    For i = 1 To q.Components(7).Quantity
        gussets(j) = q.Components(7).Name
        j = j + 1
    Next
    If pairFits Then
        For i = 0 To Int(totalGussets / 2) - 1
            items(k, 0) = gussets(i * 2) & " + " & gussets(i * 2 + 1) & " (opposed pair)"
            items(k, 1) = pairWidth
            items(k, 2) = q.GussetHeight
            items(k, 3) = (q.GussetTopWidth + q.GussetBottomWidth) * q.GussetHeight
            items(k, 4) = 2
            k = k + 1
        Next
        If (totalGussets Mod 2) = 1 Then
            items(k, 0) = gussets(totalGussets - 1)
            items(k, 1) = q.GussetBottomWidth
            items(k, 2) = q.GussetHeight
            items(k, 3) = (q.GussetTopWidth + q.GussetBottomWidth) / 2 * q.GussetHeight
            items(k, 4) = 1
            k = k + 1
        End If
    Else
        For i = 0 To totalGussets - 1
            items(k, 0) = gussets(i)
            items(k, 1) = q.GussetBottomWidth
            items(k, 2) = q.GussetHeight
            items(k, 3) = (q.GussetTopWidth + q.GussetBottomWidth) / 2 * q.GussetHeight
            items(k, 4) = 1
            k = k + 1
        Next
    End If
    baseline = count + 2
    For sortOrder = 0 To 3
        UINestSort items, sortOrder
        For splitMode = 0 To 5
            trial = UINestPack(items, 120 - 2 * q.EdgeMarginLength, 60 - 2 * q.EdgeMarginWidth, _
                               q.PartGap, 0, splitMode, sheets)
            If sheets < baseline Then baseline = sheets
        Next
        For splitMode = 0 To 1
            trial = UINestMaxRects(items, 120 - 2 * q.EdgeMarginLength, 60 - 2 * q.EdgeMarginWidth, _
                                  q.PartGap, 0, splitMode, sheets)
            If sheets < baseline Then baseline = sheets
        Next
    Next
    Assert q.SheetCount <= baseline, "Nesting is no worse than the prior 32 baseline layouts"
    Assert q.NestBaselineSheets = baseline, "Reported baseline matches independent 32 baseline evaluation"
End Sub
Sub CheckMaxRects(items, placements, sheetLength, sheetWidth, gap, margin, sheets)
    Dim i, j, dx, dy, used, area, expectedArea, seen
    Set seen = CreateObject("Scripting.Dictionary")
    area = 0
    expectedArea = 0
    Assert UBound(items) = UBound(placements), "MaxRects retains item count"
    For i = 0 To UBound(placements)
        Assert placements(i, 0) >= 1 And placements(i, 0) <= sheets, "MaxRects sheet index"
        Assert placements(i, 2) >= margin - 0.000001 And placements(i, 3) >= margin - 0.000001, "MaxRects near margins"
        Assert placements(i, 2) + placements(i, 4) <= margin + sheetLength + 0.000001, "MaxRects far length"
        Assert placements(i, 3) + placements(i, 5) <= margin + sheetWidth + 0.000001, "MaxRects far width"
        Assert placements(i, 6) = 0 Or placements(i, 6) = 90, "MaxRects supported rotation"
        Assert Not seen.Exists(placements(i, 1)), "MaxRects no duplicate items"
        seen.Add placements(i, 1), True
        area = area + placements(i, 7)
        expectedArea = expectedArea + items(i, 3)
        For j = 0 To i - 1
            If placements(i, 0) = placements(j, 0) Then
                dx = UINestMax(placements(j, 2) - placements(i, 2) - placements(i, 4), _
                              placements(i, 2) - placements(j, 2) - placements(j, 4))
                dy = UINestMax(placements(j, 3) - placements(i, 3) - placements(i, 5), _
                              placements(i, 3) - placements(j, 3) - placements(j, 5))
                Assert dx >= gap - 0.000001 Or dy >= gap - 0.000001, "MaxRects excludes occupied space and reserves clearance"
            End If
        Next
    Next
    Assert Abs(area - expectedArea) < 0.000001, "MaxRects profile area conservation"
    For i = 0 To UBound(items)
        Assert seen.Exists(items(i, 0)), "MaxRects all input identities present"
    Next
End Sub
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
Assert q.PricePerPound = 0.9, "Default price"
Assert Abs(q.TotalWeight - q.TotalArea * q.Thickness * q.Density) < 0.000001, "Profile weight formula"
Assert Abs(q.TotalMaterialCost - q.TotalWeight * 0.9) < 0.000001, "Material cost formula"
Set other = UIQuoteWithPrice(4, "N/A", "Double return", "Borcon Weathering Steel", "1/4 inch", 300, 450, 48, 1.8, 0.25, 0.25)
Assert other.Density = q.Density And other.TotalWeight = q.TotalWeight, "Same density for both materials"
Assert Abs(other.TotalMaterialCost - q.TotalMaterialCost * 2) < 0.000001, "User-entered price drives cost"
Set other = UICalculateQuote(4, "N/A", "Single return 1 inch", "Mild Steel", "3/16 inch", 400, 400, 23.75)
Assert Abs(other.TotalArea - 71582.9430666667) < 0.000001, "User example area retained"
Assert Abs(other.TotalWeight - 3806.4128265) < 0.00001, "User example weight"
Assert InStr(report, "SELL PRICE: $" & FormatNumber(q.SellPrice, 2)) > 0, "Report shows sell price"
Assert InStr(report, "NOT CALCULATED") = 0, "Purchased-sheet cost is now priced"
Assert q.HasPricing And Not q.Painted And q.MarginPercent = 50, "Default pricing applied"
Assert Abs(q.SheetWeight - q.SheetCount * 7200 * 0.25 * q.Density) < 0.000001, "Purchased sheet weight"
Assert Abs(q.SheetMaterialCost - q.SheetWeight * 0.9) < 0.000001 And q.SheetMaterialCost > q.TotalMaterialCost, "Material includes scrap"
Assert q.CutCost = 75 * q.SheetCount, "Cutting per sheet"
Assert q.WeldedParts = 6 + 6 + 8 + 8 + 2 + 2 + 28 + 50 And q.WeldCost = 10 * q.WeldedParts, "Welding per part"
Assert q.BentParts = 6 + 8 + 2 + 2 And q.BendCost = 5 * q.BentParts, "Bending only walls and side walls with returns"
Assert q.PaintCost = 0 And q.PaintArea = 0, "Unpainted has no paint cost"
Assert Abs(q.TotalCost - (q.SheetMaterialCost + q.LaborCost)) < 0.000001, "Total cost"
Assert Abs(q.SellPrice - q.TotalCost * 2) < 0.000001 And Abs(q.Profit - q.SellPrice * 0.5) < 0.000001, "Gross margin sell price"
UIQuotePricing q, True, 100, 10, 4, 3, 25
Assert Abs(q.PaintArea - 2 * q.TotalArea / 144) < 0.000001 And Abs(q.PaintCost - 3 * q.PaintArea) < 0.000001, "Paint both sides"
Assert q.CutCost = 100 * q.SheetCount And q.BendCost = 10 * q.BentParts And q.WeldCost = 4 * q.WeldedParts, "Custom labor rates"
Assert Abs(q.SellPrice - q.TotalCost / 0.75) < 0.000001, "Custom margin"
pdfText = q.TotalCost
UIQuoteSetMargin q, 40
Assert q.TotalCost = pdfText And Abs(q.SellPrice - q.TotalCost / 0.6) < 0.000001, "Margin change only moves sell price"
summary = UIQuotePricingCards(q)
Assert summary(2, 1) = "$" & FormatNumber(q.PaintCost, 2) And summary(3, 1) = "$" & FormatNumber(q.TotalCost, 2), "Pricing tiles"
Assert InStr(summary(1, 2), "bend 18 parts") > 0, "Labor tooltip breakdown"
summary = UIQuoteSummaryCards(q)
Assert summary(0, 0) = "QUOTE SELL PRICE  (40.00% margin)" And summary(0, 1) = "$" & FormatNumber(q.SellPrice, 2), "Headline sell price"
Assert InStr(summary(1, 1), "Mild Steel (painted)") > 0, "Painted shown in design"
For Each price In Array(-1, 95.5, 100)
    On Error Resume Next
    UIQuoteSetMargin q, price
    code = Err.Number
    Err.Clear
    On Error GoTo 0
    Assert code <> 0, "Out-of-range margin rejected"
Next
On Error Resume Next
UIQuotePricing q, False, -1, 5, 8, 2.5, 30
code = Err.Number
Err.Clear
On Error GoTo 0
Assert code <> 0, "Negative rate rejected"
Assert UIQuoteParseRate("0", "x") = 0 And UIQuoteParseRate("0.00", "x") = 0 And UIQuoteParseRate(" 2.5 ", "x") = 2.5, "Rates accept zero"
For Each price In Array("", "-1", "$5", "1,000", "5%", "0.0.0", ".")
    On Error Resume Next
    expected = UIQuoteParseRate(price, "Rate")
    code = Err.Number
    Err.Clear
    On Error GoTo 0
    Assert code <> 0, "Invalid rate rejected: " & price
Next
Set other = UICalculateQuote(4, "N/A", "No return", "Mild Steel", "1/4 inch", 300, 450, 48)
Assert other.BentParts = 0 And other.BendCost = 0, "No return means no bending"
UIQuotePricing q, False, 75, 5, 8, 2.5, 30
originalArea = q.TotalArea
originalWeight = q.TotalWeight
originalCost = q.TotalMaterialCost
summary = UIQuoteSummaryCards(q)
rows = UIQuoteComponentRows(q)
Assert UBound(summary, 1) = 6 And UBound(summary, 2) = 1, "Headline, design and five summary tiles"
Assert summary(0, 0) = "QUOTE SELL PRICE  (30.00% margin)" And summary(0, 1) = "$" & FormatNumber(q.SellPrice, 2), "Summary leads with sell price"
Assert InStr(summary(1, 1), "1/4 in  |  $0.90/lb") > 0, "Readable thickness and price"
Assert InStr(summary(1, 1), "300.000 x 450.000 x 48.000 in") > 0, "Rounded summary dimensions"
Assert InStr(summary(1, 1), "CENTER_WALL_PATTERN_COUNT") = 0, "Technical equations not in summary"
Assert summary(2, 1) = FormatNumber(q.TotalWeight, 2) & " lb" And summary(4, 1) = CStr(q.SheetCount), "Weight and sheet tiles"
Assert summary(5, 1) = FormatNumber(q.UtilizationPercent, 2) & "%" And summary(6, 0) = "Unused / potential scrap", "Utilization and scrap tiles"
Assert UBound(rows, 1) = 5 And UBound(rows, 2) = 5, "Six wall/base/end rows"
gussetRows = UIQuoteGussetRows(q)
totalRows = UIQuoteTotalRows(q)
Assert UBound(gussetRows, 1) = 1 And UBound(gussetRows, 2) = 6, "Two gusset rows and seven columns"
Assert gussetRows(0, 0) = "Mounting Gusset" And gussetRows(1, 0) = "Center Gusset", "Gussets appear once in separate table"
Assert gussetRows(0, 1) = "28" And gussetRows(1, 1) = "50", "Gusset quantities"
Assert gussetRows(0, 2) = "45.650" And gussetRows(0, 3) = "7.750" And gussetRows(0, 4) = "23.250", "Height top width base width ordering"
Assert gussetRows(0, 5) = FormatNumber(q.Components(6).TotalWeight, 2), "Gusset weight"
Assert gussetRows(1, 6) = "$" & FormatNumber(q.Components(7).TotalMaterialCost, 2), "Gusset cost"
Assert totalRows(0, 0) = "TOTAL MATERIAL - ALL COMPONENTS", "Combined total row label"
Assert totalRows(0, 1) = FormatNumber(q.TotalWeight, 2) & " lb", "Total row weight includes gussets"
Assert totalRows(0, 2) = "$" & FormatNumber(q.TotalMaterialCost, 2), "Total row cost from unrounded calculation"
Assert rows(2, 0) = "Side Planter Wall" And rows(2, 1) = "8", "Table name and quantity"
Assert rows(2, 2) = "112.375" And rows(2, 3) = "49.191", "Three decimal dimensions"
Assert rows(2, 4) = FormatNumber(q.Components(2).TotalWeight, 2), "Two decimal weight"
Assert rows(2, 5) = "$" & FormatNumber(q.Components(2).TotalMaterialCost, 2), "Two decimal cost"
Assert q.TotalArea = originalArea And q.TotalWeight = originalWeight And q.TotalMaterialCost = originalCost, "Presentation does not round stored results"
Set other = UICalculateQuote(1, "N/A", "No return", "Mild Steel", "1/4 inch", 100, 60, 48)
rows = UIQuoteComponentRows(other)
total = 0
For i = 0 To 5
    If other.Components(i).Quantity > 0 Then total = total + 1
Next
Assert total < 6, "Reference design has zero-quantity components"
Assert UBound(rows, 1) = total - 1, "Zero-quantity components hidden from table"
For i = 0 To UBound(rows, 1)
    Assert rows(i, 1) <> "0", "No zero-quantity rows shown"
Next
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
                If q.HasValidTotals And q.NestAvailable Then CheckNest q
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
Assert Not q.NestAvailable And q.SheetCount = 0, "No nesting on count mismatch"
summary = UIQuoteSummaryCards(q)
rows = UIQuoteComponentRows(q)
Assert InStr(summary(0, 0), "REVIEW REQUIRED") = 1 And summary(0, 1) = "Review required", "Warnings visible in summary"
Assert InStr(summary(0, 1), "$") = 0 And summary(2, 1) = "Withheld" And summary(4, 1) = "N/A", "No misleading estimate on mismatch"
Assert InStr(summary(1, 1), q.Warnings) > 0, "Summary preserves mismatch explanation"
Assert rows(0, 1) = "Review" And rows(0, 4) = "Withheld" And rows(0, 5) = "Withheld", "Table withholds untrusted totals"
gussetRows = UIQuoteGussetRows(q)
totalRows = UIQuoteTotalRows(q)
Assert totalRows(0, 1) = "Withheld" And totalRows(0, 2) = "Withheld", "Combined total withheld"
Assert gussetRows(0, 1) = "Review" And gussetRows(0, 5) = "Withheld" And gussetRows(0, 6) = "Withheld", "Gusset totals withheld"
Assert rows(0, 2) = "60.250", "Per-part dimensions still available"
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
    Set q = UIQuoteWithPrice(1, "N/A", "No return", "Mild Steel", "1/4 inch", 100, 100, 10, price, 0.25, 0.25)
    code = Err.Number
    Err.Clear
    On Error GoTo 0
    Assert code <> 0, "Nonpositive price rejected"
Next
Set q = UICalculateQuote(4, "N/A", "Double return", "Mild Steel", "1/4 inch", 300, 200, 20)
CheckNest q
Assert q.SheetCount = 7 And q.AreaMinimumSheets = 6, "Reference heuristic sheets versus area lower bound"
UINestQuote q, 0.25, 0.25
Assert Abs(q.UnusedPercent - 27.3066293650793) < 0.000001, "Reference scrap estimate"
pdfText = UIPDFOutline(Array(Array(0, 0), Array(10, 0), Array(3, 20), Array(0, 20)), 5, 7, 90, 20, True)
Assert InStr(pdfText, "186 232 m") > 0, "Rotated gusset origin"
Assert InStr(pdfText, "66 250 l") > 0, "Rotated gusset top vertex"
pdfText = UIPDFOutline(Array(Array(0, 0), Array(10, 0), Array(3, 20), Array(0, 20)), 5, 7, 180, 20, True)
Assert InStr(pdfText, "126 352 m") > 0, "180-degree gusset origin"
pdfText = UIPDFOutline(Array(Array(0, 0), Array(10, 0), Array(3, 20), Array(0, 20)), 5, 7, 270, 20, True)
Assert InStr(pdfText, "66 292 m") > 0, "270-degree gusset origin"
Assert InStr(UIPDFText(0, 0, 10, "A (B) \ C"), "A \(B\) \\ C") > 0, "PDF text escapes parentheses and slash"
Assert UIPDFPartLabel(0) = "CW" And UIPDFPartLabel(1) = "CBP" And UIPDFPartLabel(2) = "SW", "Planter part labels"
Assert UIPDFPartLabel(3) = "SBP" And UIPDFPartLabel(4) = "LW" And UIPDFPartLabel(5) = "RW", "Base plate and end wall labels"
Assert UIPDFPartLabel(6) = "MG" And UIPDFPartLabel(7) = "CG", "Gusset labels"
Assert q.Components(4).Name = "Left Side Wall" And q.Components(5).Name = "Right Side Wall", "Readable end wall names"
Set other = UICalculateQuote(1, "N/A", "No return", "Mild Steel", "3/16 inch", 100, 100, 10)
pdfText = UIPDFParts(other)
Assert InStr(pdfText, "(Center Planter Wall)") > 0 And InStr(pdfText, "(Mounting Gusset)") > 0, "Used parts listed"
Assert InStr(pdfText, "Side Planter") = 0 And InStr(pdfText, "Side Wall)") = 0, "Unused parts hidden from parts list"
pdfText = UIQuoteLayoutPDF(q)
Set pdfFSO = CreateObject("Scripting.FileSystemObject")
UISaveQuoteLayoutPDF q, WScript.Arguments(0)
Set pdfFile = pdfFSO.OpenTextFile(WScript.Arguments(0), 1)
Assert pdfFile.ReadAll = pdfText, "File writer preserves PDF byte offsets"
pdfFile.Close
On Error Resume Next
UISaveQuoteLayoutPDF q, WScript.Arguments(0)
code = Err.Number
Err.Clear
On Error GoTo 0
Assert code <> 0, "Existing PDF never overwritten"
UIPreviewQuoteLayoutPDF q
Assert pdfFSO.FileExists(previewPath), "Temporary PDF exists before launching viewer"
Assert pdfFSO.GetParentFolderName(previewPath) = CStr(pdfFSO.GetSpecialFolder(2)), "Preview uses system temporary folder"
pdfFSO.DeleteFile previewPath
previewPath = ""
Set other = UICalculateQuote(1, "N/A", "No return", "Mild Steel", "1/4 inch", 10, 10, 6)
WScript.Echo "Example: " & q.SheetCount & " sheets; unused=" & q.UnusedPercent
p = q.SheetPlacements
UINestQuote q, 0.25, 0.25
Assert UBound(q.SheetPlacements) = UBound(p), "Deterministic repeat layout"
For i = 0 To UBound(p)
    For j = 0 To 8
        Assert p(i,j) = q.SheetPlacements(i,j), "Deterministic placement fields"
    Next
Next
Set q = UICalculateQuote(1, "N/A", "No return", "Mild Steel", "1/4 inch", 120, 100, 20)
WScript.Echo "Relaxed: " & q.NestWarning
Assert q.NestAvailable, "120-inch part nests by using the sheet ends"
Assert q.EdgeMarginLength = 0 And q.EdgeMarginWidth = 0.25, "Only the length-direction margin is reduced"
Assert q.EdgeMargin = 0.25, "Requested margin is preserved"
Assert InStr(q.NestWarning, "sheet ends") > 0 And InStr(q.NestWarning, "sheet sides") = 0, "Warning names the reduced edge"
CheckNest q
Assert q.HasPricing And q.SellPrice > 0, "Relaxed-margin layout is priced"
Assert InStr(UIQuoteReport(q), "WARNING: Edge margin reduced") > 0, "Report shows reduced-margin warning"
pdfText = UIQuoteLayoutPDF(q)
Assert InStr(pdfText, "NOTE: Edge margin reduced") > 0, "PDF summary shows reduced-margin warning"
relaxLength = 0.25
relaxWidth = 0.25
relaxNames = ""
Assert UINestRelaxMargin("Fits", 119, 59, 0.25, relaxLength, relaxWidth, relaxNames), "Part within margins fits"
Assert relaxLength = 0.25 And relaxWidth = 0.25 And relaxNames = "", "No relaxation when margins fit"
Assert UINestRelaxMargin("Wide", 100, 59.8, 0.25, relaxLength, relaxWidth, relaxNames), "Full-width part fits"
Assert relaxLength = 0.25 And Abs(relaxWidth - 0.1) < 0.000001 And relaxNames = "Wide", "Only the width margin is reduced"
Assert Not UINestRelaxMargin("Too big", 120, 61, 0.25, relaxLength, relaxWidth, relaxNames), "Oversize part rejected"
Set q = UICalculateQuote(4, "N/A", "Double return", "Mild Steel", "3/16 inch", 400, 205, 40)
WScript.Echo "400x205x40: sheets=" & q.SheetCount & " sell=" & q.SellPrice & " warn=" & q.NestWarning
Assert q.NestAvailable And q.HasPricing And q.EdgeMarginLength = 0, "120-inch center walls are priced"
CheckNest q
q.NestAvailable = False
On Error Resume Next
pdfText = UIQuoteLayoutPDF(q)
code = Err.Number
Err.Clear
On Error GoTo 0
Assert code <> 0, "Unavailable layout cannot preview PDF"
On Error Resume Next
UIPreviewQuoteLayoutPDF q
code = Err.Number
Err.Clear
On Error GoTo 0
Assert code <> 0 And previewPath = "", "Unavailable preview never launches viewer"
UINestQuote q, 0.25, 0
CheckNest q
Set q = UICalculateQuote(1, "N/A", "No return", "Mild Steel", "1/4 inch", 100, 100, 20)
pdfText = UIQuoteLayoutPDF(q)
Assert InStr(pdfText, "ZERO GAP:") = 0, "Nonzero gap has no co-edge warning"
Assert InStr(pdfText, "CypCut can use finer angles") > 0, "PDF explains production rotation difference"
UINestQuote q, 0, 0
CheckNest q
pdfText = UIQuoteLayoutPDF(q)
Assert InStr(pdfText, "ZERO GAP: touching edges only") > 0, "Zero-gap PDF warns about co-edge setup"
Assert InStr(pdfText, "Operations must set up co-edge/common-line cuts in CypCut") > 0, "Zero gap is not a toolpath"
customItems(0, 0) = "Rotation test"
customItems(0, 1) = 40
customItems(0, 2) = 70
customItems(0, 3) = 2800
customItems(0, 4) = 1
p = UINestPack(customItems, 119.5, 59.5, 0.25, 0.25, 0, customSheets)
Assert customSheets = 1 And p(0, 6) = 90, "Part requiring 90 degree rotation fits"
Assert p(0, 4) = 70 And p(0, 5) = 40, "Rotated envelope dimensions"
p = UINestMaxRects(customItems, 119.5, 59.5, 0.25, 0.25, 0, customSheets)
Assert customSheets = 1 And p(0, 6) = 90, "MaxRects required rotation fits"
CheckMaxRects customItems, p, 119.5, 59.5, 0.25, 0.25, customSheets
trialWidths = Array(50, 40, 50, 30, 25, 25, 40, 40)
trialHeights = Array(15, 20, 20, 10, 15, 25, 25, 35)
For i = 0 To 7
    heuristicItems(i, 0) = "Synthetic " & i
    heuristicItems(i, 1) = trialWidths(i)
    heuristicItems(i, 2) = trialHeights(i)
    heuristicItems(i, 3) = trialWidths(i) * trialHeights(i)
    heuristicItems(i, 4) = 1
Next
legacyBest = 100
For sortOrder = 0 To 1
    ReDim trialItems(7, 4)
    For i = 0 To 7
        For j = 0 To 4
            trialItems(i, j) = heuristicItems(i, j)
        Next
    Next
    UINestSort trialItems, sortOrder
    For splitMode = 0 To 1
        p = UINestPack(trialItems, 60, 40, 0, 0, splitMode, trialSheets)
        If trialSheets < legacyBest Then legacyBest = trialSheets
    Next
Next
extraBest = legacyBest
For sortOrder = 0 To 3
    ReDim trialItems(7, 4)
    For i = 0 To 7
        For j = 0 To 4
            trialItems(i, j) = heuristicItems(i, j)
        Next
    Next
    UINestSort trialItems, sortOrder
    For splitMode = 0 To 5
        If sortOrder >= 2 Or splitMode >= 2 Then
            p = UINestPack(trialItems, 60, 40, 0, 0, splitMode, trialSheets)
            If trialSheets < extraBest Then extraBest = trialSheets
        End If
    Next
Next
Assert legacyBest = 4 And extraBest = 3, "Additional heuristic trials improve synthetic fit from four sheets to three"
trialWidths = Array(15, 45, 35, 15, 10, 40, 20, 20, 50, 20)
trialHeights = Array(35, 20, 15, 25, 25, 15, 20, 25, 15, 15)
For i = 0 To 9
    maxItems(i, 0) = "MaxRects regression " & i
    maxItems(i, 1) = trialWidths(i)
    maxItems(i, 2) = trialHeights(i)
    maxItems(i, 3) = trialWidths(i) * trialHeights(i)
    maxItems(i, 4) = 1
Next
extraBest = 100
maxBest = 100
For sortOrder = 0 To 3
    trialItems = maxItems
    UINestSort trialItems, sortOrder
    For splitMode = 0 To 5
        p = UINestPack(trialItems, 60, 40, 0.25, 0.25, splitMode, trialSheets)
        If trialSheets < extraBest Then extraBest = trialSheets
    Next
    For splitMode = 0 To 1
        p = UINestMaxRects(trialItems, 60, 40, 0.25, 0.25, splitMode, trialSheets)
        CheckMaxRects trialItems, p, 60, 40, 0.25, 0.25, trialSheets
        If trialSheets < maxBest Then maxBest = trialSheets
    Next
Next
Assert extraBest = 4 And maxBest = 3, "MaxRects improves packing versus all24 split layouts from four sheets to three"
For i = 0 To 95
    stressItems(i, 0) = "Stress " & i
    stressItems(i, 1) = 8 + (i Mod 7)
    stressItems(i, 2) = 6 + (i Mod 11)
    stressItems(i, 3) = stressItems(i, 1) * stressItems(i, 2)
    stressItems(i, 4) = 1
Next
For Each gap In Array(0, 0.25, 1)
    For splitMode = 0 To 1
        p = UINestMaxRects(stressItems, 119.5, 59.5, gap, 0.25, splitMode, trialSheets)
        CheckMaxRects stressItems, p, 119.5, 59.5, gap, 0.25, trialSheets
    Next
Next
Set other = UICalculateQuote(1, "N/A", "No return", "Mild Steel", "1/4 inch", 100, 100, 20)
For i = 0 To UBound(other.Components)
    other.Components(i).Quantity = 0
Next
other.Components(6).Quantity = 3
other.TotalArea = other.Components(6).Area * 3
UINestQuote other, 0.25, 25
Assert other.NestAvailable And UBound(other.SheetPlacements) = 2, "Oversized pairs fall back to three fitting single gussets"
For i = 0 To UBound(other.SheetPlacements)
    Assert other.SheetPlacements(i, 8) = 1, "Fallback preserves physical single counts"
    Assert other.SheetPlacements(i, 2) + other.SheetPlacements(i, 4) <= 95.000001, "Fallback length boundary"
    Assert other.SheetPlacements(i, 3) + other.SheetPlacements(i, 5) <= 35.000001, "Fallback width boundary"
Next
For Each margin In Array(-1, 30)
    On Error Resume Next
    UINestQuote q, 0.25, margin
    code = Err.Number
    Err.Clear
    On Error GoTo 0
    Assert code <> 0, "Invalid margin explicitly rejected"
Next
On Error Resume Next
UINestQuote q, -0.01, 0.25
code = Err.Number
Err.Clear
On Error GoTo 0
Assert code <> 0, "Negative gap rejected"
q.Components(0).Quantity = 2001
UINestQuote q, 0.25, 0.25
Assert Not q.NestAvailable And q.SheetCount = 0 And InStr(q.NestStatus, "2000") > 0, "Bounded layout unavailable, not false success"
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
            layout = Path(folder) / "layout.pdf"
            runner.write_text(script, encoding="ascii")
            result = subprocess.run(["cscript", "//Nologo", str(runner), str(layout)], capture_output=True,
                                    text=True, timeout=120)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertIn("PASS:", result.stdout, result.stdout + result.stderr)
            from pypdf import PdfReader
            reader = PdfReader(layout, strict=True)
            self.assertEqual(len(reader.pages), 10)
            summary_text = reader.pages[0].extract_text()
            self.assertIn("Optimization Summary", summary_text)
            self.assertIn("SHEETS REQUIRED", summary_text)
            self.assertIn("78 physical parts", summary_text)
            self.assertIn("Used for profiles: 72.69%", summary_text)
            self.assertIn("Unused / potential scrap: 27.31%", summary_text)
            self.assertIn("350.00 sq ft", summary_text)
            self.assertIn("7 sheets of 120 x 60 in", summary_text)
            self.assertIn("Profile area: 254.43 sq ft", summary_text)
            self.assertIn("Unused area: 95.57 sq ft", summary_text)
            self.assertIn("Area-only minimum: 6 sheets", summary_text)
            self.assertIn("(27 pairs)", summary_text)
            self.assertIn("Guillotine - paired gussets", summary_text)
            self.assertIn("32-trial baseline", summary_text)
            self.assertIn("Up to 48 different layouts are tried", summary_text)
            self.assertIn("(4 parts turned)", summary_text)
            self.assertIn("4. Remove a sheet if possible", summary_text)
            for jargon in ("MaxRects packing", "Guillotine and", "least-used sheet"):
                self.assertNotIn(jargon, summary_text)
            self.assertNotIn("Compare 48", summary_text)
            self.assertIn("Sheet Usage - Unused / Potential Scrap", reader.pages[1].extract_text())
            parts_text = reader.pages[2].extract_text()
            self.assertIn("Parts List - Total Quantities", parts_text)
            for label, name, quantity in (("CW", "Center Planter Wall", 6), ("CBP", "Center Planter Base Plate", 6),
                                          ("SW", "Side Planter Wall", 4), ("SBP", "Side Planter Base Plate", 4),
                                          ("LW", "Left Side Wall", 2), ("RW", "Right Side Wall", 2),
                                          ("MG", "Mounting Gusset", 20), ("CG", "Center Gusset", 34)):
                self.assertRegex(parts_text, rf"{label}\s*{name}\s*(Rectangle|Trapezoid)[^\n]*?\s{quantity}\s")
            self.assertRegex(parts_text, r"TOTAL\s*78\s*254\.43")
            self.assertIn("not a guaranteed global minimum", summary_text)
            self.assertIn("not all discarded scrap", summary_text)
            self.assertIn("0.25 in / 0.25 in", summary_text)
            summary_stream = reader.pages[0].get_contents().get_data().decode("ascii")
            self.assertIn("re f Q", summary_stream)
            self.assertIn("/F2 22 Tf", summary_stream)
            area = 0
            parts = 0
            legend_counts = [0] * 8
            for sheet, page in enumerate(reader.pages[3:], 1):
                self.assertEqual(list(page.mediabox), [0, 0, 792, 612])
                text = page.extract_text()
                self.assertIn(f"Sheet {sheet} of 7", text)
                self.assertIn("NOT A PRODUCTION CUTTING PLAN", text)
                self.assertIn("Stock: 120 x 60 in", text)
                polygon = []
                polygons = []
                page_area = 0
                page_parts = 0
                for line in page.get_contents().get_data().decode("ascii").splitlines():
                    tokens = line.split()
                    if len(tokens) == 3 and tokens[-1] in ("m", "l"):
                        polygon.append(((float(tokens[0]) - 36) / 6,
                                        (float(tokens[1]) - 190) / 6))
                    elif line == "h S":
                        polygon = []
                    elif line == "h B":
                        self.assertEqual(len(polygon), 4)
                        signed = 0
                        for index, (x, y) in enumerate(polygon):
                            xx, yy = polygon[(index + 1) % 4]
                            self.assertGreaterEqual(x, 0.25 - 1e-8)
                            self.assertLessEqual(x, 119.75 + 1e-8)
                            self.assertGreaterEqual(y, 0.25 - 1e-8)
                            self.assertLessEqual(y, 59.75 + 1e-8)
                            signed += x * yy - xx * y
                        page_area += abs(signed) / 2
                        page_parts += 1
                        polygons.append(polygon)
                        polygon = []
                for i, first in enumerate(polygons):
                    for second in polygons[i + 1:]:
                        self.assertFalse(polygons_overlap(first, second),
                                         f"Overlapping component profiles on sheet {sheet}")
                        self.assertGreaterEqual(polygon_clearance(first, second), 0.25 - 1e-8,
                                                f"Insufficient part clearance on sheet {sheet}")
                area += page_area
                parts += page_parts
                self.assertIn(f"This sheet: {page_parts} parts", text)
                self.assertIn(f"Utilization: {page_area / 7200 * 100:.2f}%", text)
                self.assertIn(f"Unused: {100 - page_area / 7200 * 100:.2f}%", text)
                names = ("Center Planter Wall", "Center Planter Base Plate", "Side Planter Wall",
                         "Side Planter Base Plate", "Left Side Wall", "Right Side Wall",
                         "Mounting Gusset", "Center Gusset")
                labels = ("CW", "CBP", "SW", "SBP", "LW", "RW", "MG", "CG")
                self.assertIn("PART LEGEND", text)
                self.assertNotIn("C1", text)
                quantities = []
                for code, (name, label) in enumerate(zip(names, labels), 1):
                    match = re.search(rf"\b{label}\s*{re.escape(name)}: (\d+)", text)
                    if match is None:
                        self.assertNotIn(name + ":", text)
                        continue
                    self.assertGreater(int(match[1]), 0, "Legend lists only parts on this sheet")
                    quantities.append((str(code), match[1]))
                self.assertEqual(sum(int(q) for _, q in quantities), page_parts)
                for code, quantity in quantities:
                    legend_counts[int(code) - 1] += int(quantity)
            self.assertEqual(parts, 78)
            self.assertEqual(legend_counts, [6, 6, 4, 4, 2, 2, 20, 34])
            self.assertAlmostEqual(area, 36637.4588, places=6)

    @unittest.skipUnless(shutil.which("cscript"), "Windows Script Host required")
    def test_demonstrable_improvements_and_sheet_elimination(self):
        """Verify synthetic improvements for single gussets, mixed gussets, and sheet elimination."""
        def procedure(source, name):
            match = re.search(rf"^(?:Public|Private) (Function|Sub) {name}\b.*?^End \1$",
                              source, re.MULTILINE | re.DOTALL)
            body = re.sub(r" As (?:Long|String|Double|Boolean|Variant|Object|UIQuoteResult|UIQuoteComponent)\b", "", match[0])
            body = body.replace("Private Function", "Function").replace("Public Function", "Function")
            body = body.replace("Private Sub", "Sub").replace("Public Sub", "Sub")
            body = body.replace("parts(0 To 7)", "parts(7)")
            body = body.replace("rows(0 To count - 1, 0 To weightColumn + 1)", "rows(count - 1, weightColumn + 1)")
            body = body.replace("rows(0 To 0, 0 To 2)", "rows(0, 2)")
            body = re.sub(r"(\w+)\(0 To ([^\n,]+), 0 To ([^\n)]+)\)", r"\1(\2, \3)", body)
            body = re.sub(r"(\w+)\(0 To ([^\n)]+)\)", r"\1(\2)", body)
            body = body.replace("usedSheets(1 To sheets)", "usedSheets(sheets)").replace("0#", "0")
            body = re.sub(r"^(\s*Next) \w+$", r"\1", body, flags=re.MULTILINE)
            body = re.sub(r"\bTrim\$", "Trim", body).replace("ByRef gussets()", "ByRef gussets")
            body = body.replace("Right$(", "Right(").replace("Format$(", "FormatNumberPDF(")
            body = body.replace("Left$(", "Left(").replace("ByRef offsets()", "ByRef offsets").replace("Mid$(", "Mid(").replace("Trim$(", "Trim(")
            body = body.replace("Str$(value)", "CStr(value)").replace("Val(cleaned)", "CDbl(cleaned)")
            return body

        fields = re.sub(r" As (?:Long|String|Double|Boolean|Variant)\b", "",
                        RESULT.split("Option Explicit", 1)[1])
        base_script = "Option Explicit\nSetLocale 1033\n"
        base_script += "Sub Assert(cond, msg)\n    If Not cond Then\n        WScript.Echo \"FAIL: \" & msg\n        WScript.Quit 1\n    End If\nEnd Sub\n"
        base_script += "Function IIf(cond, a, b)\n    If cond Then IIf = a Else IIf = b\nEnd Function\n"
        base_script += "Function FormatNumberPDF(value, pattern)\n    FormatNumberPDF = FormatNumber(value, 2, -1, 0, 0)\nEnd Function\n"
        base_script += "\nClass UIQuoteResult\n" + fields + "\nEnd Class\n"
        comp_fields = re.sub(r" As (?:String|Double)\b", "", COMPONENT.split("Option Explicit", 1)[1])
        base_script += "\nClass UIQuoteComponent\n" + comp_fields + "\nEnd Class\n"
        for name in ("UINestQuote", "UINestSort", "UINestMax", "UINestPack", "UINestMaxRects",
                     "UINestFreeAdd", "UINestFreePrune", "UINestReport", "UINestBuildItems",
                     "UINestFreeSubtract", "UINestEliminateSheet", "UINestRelaxMargin", "UINestMarginText"):
            base_script += procedure(NESTING, name) + "\n"

        test_body = base_script + r"""
Dim q, i, comps(7), p(1, 8), outP, ok, j, orig(1, 8), p3(2, 8), orig3(2, 8)

' 1. Single-gusset demonstrable improvement: 2 baseline sheets reduced to 1 sheet
Set q = New UIQuoteResult
q.Material = "Mild Steel"
q.Thickness = 0.25
q.HasValidTotals = True
q.GussetBottomWidth = 20
q.GussetTopWidth = 5
q.GussetHeight = 25
For i = 0 To 7
    Set comps(i) = New UIQuoteComponent
    comps(i).Material = q.Material
    comps(i).Thickness = q.Thickness
    comps(i).Quantity = 0
Next
comps(0).Name = "Panel"
comps(0).Length = 95
comps(0).Width = 25
comps(0).Area = 95 * 25
comps(0).Quantity = 2
comps(6).Name = "Mounting Gusset"
comps(6).Length = 25
comps(6).Width = 20
comps(6).TopWidth = 5
comps(6).Area = (20 + 5) * 25 / 2
comps(6).Quantity = 2
q.Components = comps
q.TotalArea = 2 * (95 * 25) + 2 * comps(6).Area
UINestQuote q, 0.25, 0.25
Assert q.SheetCount = 1, "Single gusset policy reduces sheet count to 1"
Assert q.NestBaselineSheets = 2, "Paired baseline required 2 sheets"
Assert q.NestMethod = "MaxRects - single gussets", "Winning method is MaxRects single gussets"
Assert q.NestTrialCount = 48, "48 candidate layouts evaluated"
Assert UBound(q.SheetPlacements) = 3, "All 4 physical components placed"
For i = 0 To UBound(q.SheetPlacements)
    Assert q.SheetPlacements(i, 0) = 1, "All parts on sheet 1"
    Assert q.SheetPlacements(i, 8) = 1, "Single gussets placed as individual parts"
Next

' 2. Mixed-gusset demonstrable improvement: 2 baseline sheets reduced to 1 sheet
Set q = New UIQuoteResult
q.Material = "Mild Steel"
q.Thickness = 0.25
q.HasValidTotals = True
q.GussetBottomWidth = 16
q.GussetTopWidth = 5
q.GussetHeight = 30
For i = 0 To 7
    Set comps(i) = New UIQuoteComponent
    comps(i).Material = q.Material
    comps(i).Thickness = q.Thickness
    comps(i).Quantity = 0
Next
comps(0).Name = "Panel"
comps(0).Length = 90
comps(0).Width = 20
comps(0).Area = 90 * 20
comps(0).Quantity = 2
comps(6).Name = "Mounting Gusset"
comps(6).Length = 30
comps(6).Width = 16
comps(6).TopWidth = 5
comps(6).Area = (16 + 5) * 30 / 2
comps(6).Quantity = 5
q.Components = comps
q.TotalArea = 2 * (90 * 20) + 5 * comps(6).Area
UINestQuote q, 0.25, 0.25
Assert q.SheetCount = 1, "Mixed gusset policy reduces sheet count to 1"
Assert q.NestBaselineSheets = 2, "Paired baseline required 2 sheets"
Assert q.NestMethod = "MaxRects - mixed gussets", "Winning method is MaxRects mixed gussets"
Assert q.NestTrialCount = 48, "48 candidate layouts evaluated"

' 3. Successful Last-Sheet Elimination Pass
p(0, 0) = 1
p(0, 1) = "Part 1"
p(0, 2) = 0.25
p(0, 3) = 0.25
p(0, 4) = 80
p(0, 5) = 50
p(0, 6) = 0
p(0, 7) = 4000
p(0, 8) = 1

p(1, 0) = 2
p(1, 1) = "Part 2"
p(1, 2) = 0.25
p(1, 3) = 0.25
p(1, 4) = 30
p(1, 5) = 40
p(1, 6) = 0
p(1, 7) = 1200
p(1, 8) = 1

ok = UINestEliminateSheet(p, 2, 119.5, 59.5, 0.25, 0.25, outP)
Assert ok, "Elimination succeeded"
Assert outP(0, 0) = 1 And outP(1, 0) = 1, "All parts relocated onto Sheet 1"
Assert outP(1, 2) >= 80.25, "Relocated part respects gap clearance on Sheet 1"

' 4. Failed Elimination Pass: Source Unmutated Guarantee
p(0, 4) = 110
p(0, 5) = 55
p(1, 4) = 50
p(1, 5) = 50
For i = 0 To 1
    For j = 0 To 8
        orig(i, j) = p(i, j)
    Next
Next
ok = UINestEliminateSheet(p, 2, 119.5, 59.5, 0.25, 0.25, outP)
Assert Not ok, "Impossible elimination rejected"
For i = 0 To 1
    For j = 0 To 8
        Assert p(i, j) = orig(i, j), "Source placements strictly unmutated on failed elimination"
    Next
Next

' 5. Non-last least-used sheet: middle sheet eliminated, later sheet renumbered
For i = 0 To 2
    p3(i, 2) = 0.25
    p3(i, 3) = 0.25
    p3(i, 6) = 0
    p3(i, 8) = 1
    p3(i, 0) = i + 1
    p3(i, 1) = "Part " & (i + 1)
Next
p3(0, 4) = 80
p3(0, 5) = 50
p3(0, 7) = 4000
p3(1, 4) = 30
p3(1, 5) = 40
p3(1, 7) = 1200
p3(2, 4) = 100
p3(2, 5) = 50
p3(2, 7) = 5000
For i = 0 To 2
    For j = 0 To 8
        orig3(i, j) = p3(i, j)
    Next
Next
ok = UINestEliminateSheet(p3, 3, 119.5, 59.5, 0.25, 0.25, outP)
Assert ok, "Middle least-used sheet eliminated"
Assert outP(0, 0) = 1, "Sheet 1 part stays on sheet 1"
Assert outP(1, 0) = 1, "Middle-sheet part relocated beside sheet 1 part"
Assert outP(1, 2) >= 80.25 - 0.000001, "Relocated part keeps gap from sheet 1 part"
Assert outP(2, 0) = 2, "Former sheet 3 renumbered to sheet 2"
For j = 1 To 8
    Assert outP(2, j) = p3(2, j), "Renumbered sheet placement otherwise unchanged " & j
Next
For i = 0 To 2
    For j = 0 To 8
        Assert p3(i, j) = orig3(i, j), "Source unmutated on successful non-last elimination"
    Next
Next

' 6. Non-last failure: middle part too large for remaining space leaves source untouched
p3(1, 4) = 60
p3(1, 5) = 45
p3(1, 7) = 2700
For i = 0 To 2
    For j = 0 To 8
        orig3(i, j) = p3(i, j)
    Next
Next
ok = UINestEliminateSheet(p3, 3, 119.5, 59.5, 0.25, 0.25, outP)
Assert Not ok, "Unplaceable middle sheet rejected"
For i = 0 To 2
    For j = 0 To 8
        Assert p3(i, j) = orig3(i, j), "Source unmutated on failed non-last elimination"
    Next
Next
WScript.Echo "PASS: demonstrable improvements and sheet elimination"
"""
        with tempfile.TemporaryDirectory(prefix="test-vbs-") as folder:
            runner = Path(folder) / "run.vbs"
            runner.write_text(test_body, encoding="ascii")
            res = subprocess.run(["cscript", "//Nologo", str(runner)], capture_output=True, text=True, timeout=30)
            self.assertEqual(res.returncode, 0, res.stdout + res.stderr)
            self.assertIn("PASS: demonstrable improvements and sheet elimination", res.stdout)

    @unittest.skipUnless(shutil.which("cscript"), "Windows Script Host required")
    def test_five_realistic_reference_examples_and_determinism(self):
        """Evaluate 5 distinct realistic configurations verifying non-regression and determinism."""
        def procedure(source, name):
            match = re.search(rf"^(?:Public|Private) (Function|Sub) {name}\b.*?^End \1$",
                              source, re.MULTILINE | re.DOTALL)
            body = re.sub(r" As (?:Long|String|Double|Boolean|Variant|Object|UIQuoteResult|UIQuoteComponent)\b", "", match[0])
            body = body.replace("Private Function", "Function").replace("Public Function", "Function")
            body = body.replace("Private Sub", "Sub").replace("Public Sub", "Sub")
            body = body.replace("parts(0 To 7)", "parts(7)")
            body = body.replace("rows(0 To count - 1, 0 To weightColumn + 1)", "rows(count - 1, weightColumn + 1)")
            body = body.replace("rows(0 To 0, 0 To 2)", "rows(0, 2)")
            body = re.sub(r"(\w+)\(0 To ([^\n,]+), 0 To ([^\n)]+)\)", r"\1(\2, \3)", body)
            body = re.sub(r"(\w+)\(0 To ([^\n)]+)\)", r"\1(\2)", body)
            body = body.replace("usedSheets(1 To sheets)", "usedSheets(sheets)").replace("0#", "0")
            if name == "UICalculateQuote":
                body = body.replace("Optional ByVal pricePerPound = 0.9", "ByVal pricePerPound")
                body = body.replace("Optional ByVal partGap = 0.25", "ByVal partGap")
                body = body.replace("Optional ByVal edgeMargin = 0.25", "ByVal edgeMargin")
                body = body.replace("UICalculateQuote", "UIQuoteWithPrice")
            body = re.sub(r"^(\s*Next) \w+$", r"\1", body, flags=re.MULTILINE)
            body = re.sub(r"\bTrim\$", "Trim", body).replace("ByRef gussets()", "ByRef gussets")
            body = body.replace("Right$(", "Right(").replace("Format$(", "FormatNumberPDF(")
            body = body.replace("Left$(", "Left(").replace("ByRef offsets()", "ByRef offsets").replace("Mid$(", "Mid(").replace("Trim$(", "Trim(")
            body = body.replace("Str$(value)", "CStr(value)").replace("Val(cleaned)", "CDbl(cleaned)")
            if name in ("UIQuoteParsePrice", "UIQuoteParseRate"):
                label = "InvalidPrice" if name == "UIQuoteParsePrice" else "InvalidRate"
                body = body.replace(f"On Error GoTo {label}", "On Error Resume Next\n    Err.Clear")
                body = body.replace(f"    Exit Function\n{label}:", "    If Err.Number = 0 Then Exit Function\n    On Error GoTo 0")
            if name == "UIQuoteGussets":
                body = body.replace("    Select Case length\n", "")
                body = body.replace("Case Is <= 24:", "If length <= 24 Then\n")
                for limit in (48, 72, 96):
                    body = body.replace(f"Case Is <= {limit}:", f"ElseIf length <= {limit} Then\n")
                body = body.replace("Case Else:", "Else\n").replace("End Select", "End If")
            return body

        fields = re.sub(r" As (?:Long|String|Double|Boolean|Variant)\b", "",
                        RESULT.split("Option Explicit", 1)[1])
        base = "Option Explicit\nSetLocale 1033\n"
        base += "Sub Assert(cond, msg)\n    If Not cond Then\n        WScript.Echo \"FAIL: \" & msg\n        WScript.Quit 1\n    End If\nEnd Sub\n"
        base += "Function IIf(cond, a, b)\n    If cond Then IIf = a Else IIf = b\nEnd Function\n"
        base += "Function FormatNumberPDF(value, pattern)\n    FormatNumberPDF = FormatNumber(value, 2, -1, 0, 0)\nEnd Function\n"
        base += "\nClass UIQuoteResult\n" + fields + "\nEnd Class\n"
        comp_fields = re.sub(r" As (?:String|Double)\b", "", COMPONENT.split("Option Explicit", 1)[1])
        base += "\nClass UIQuoteComponent\n" + comp_fields + "\nEnd Class\n"
        for name in ("UICalculateQuote", "UIQuoteFlatHeight", "UIQuotePart",
                     "UIQuoteComponents", "UIQuoteGussets", "UIQuoteReport",
                     "UIQuoteSummaryCards", "UIQuoteComponentRows", "UIQuoteGussetRows",
                     "UIQuoteTableRows", "UIQuoteTotalRows", "UIQuoteSheetStats", "UIQuoteSheetRows",
                     "UIQuoteDefaultRate", "UIQuotePricing", "UIQuoteSetMargin", "UIQuoteParseRate",
                     "UIQuotePricingCards", "UIQuotePricingReport"):
            base += procedure(CALCULATOR, name) + "\n"
        for name in ("UIValidateSelections", "UIValidateDimensions", "UIInchesText", "UIParseInches"):
            base += procedure(MODULE, name) + "\n"
        base += procedure(CALCULATOR, "UIQuoteParsePrice") + "\n"
        for name in ("UINestQuote", "UINestSort", "UINestMax", "UINestPack", "UINestMaxRects",
                     "UINestFreeAdd", "UINestFreePrune", "UINestReport", "UINestBuildItems",
                     "UINestFreeSubtract", "UINestEliminateSheet", "UINestRelaxMargin", "UINestMarginText"):
            base += procedure(NESTING, name) + "\n"

        test_body = base + r"""
Function UICalculateQuote(walls, orient, ret, material, thickness, length, width, height)
    Set UICalculateQuote = UIQuoteWithPrice(walls, orient, ret, material, thickness, length, width, height, 0.9, 0.25, 0.25)
End Function

Sub CheckExample(q, name, expSheets, expBaseline, expParts, expAreaSqFt)
    Dim p, i, j, area, parts, gap, margin
    Assert q.NestAvailable, name & " nesting available"
    Assert q.SheetCount = expSheets, name & " sheet count"
    Assert q.NestBaselineSheets = expBaseline, name & " baseline sheets"
    Assert q.NestTrialCount = 48, name & " 48 candidate layouts"
    Assert Len(q.NestMethod) > 0, name & " nest method"
    Assert q.SheetCount <= q.NestBaselineSheets, name & " no regression against baseline"
    Assert Abs(Round(q.TotalArea / 144, 2) - expAreaSqFt) < 0.01, name & " total area sq ft"

    ' Check bounds, clearances, area conservation
    p = q.SheetPlacements
    gap = q.PartGap
    margin = q.EdgeMargin
    area = 0
    parts = 0
    For i = 0 To UBound(p)
        Assert p(i, 0) >= 1 And p(i, 0) <= q.SheetCount, name & " sheet index"
        Assert p(i, 2) >= q.EdgeMarginLength - 0.000001 And p(i, 3) >= q.EdgeMarginWidth - 0.000001, name & " near margin"
        Assert p(i, 2) + p(i, 4) <= 120 - q.EdgeMarginLength + 0.000001, name & " far length"
        Assert p(i, 3) + p(i, 5) <= 60 - q.EdgeMarginWidth + 0.000001, name & " far width"
        Assert p(i, 6) = 0 Or p(i, 6) = 90, name & " rotation"
        area = area + p(i, 7)
        parts = parts + p(i, 8)
        For j = i + 1 To UBound(p)
            If p(i, 0) = p(j, 0) Then
                Assert p(i, 2) + p(i, 4) + gap <= p(j, 2) + 0.000001 Or _
                       p(j, 2) + p(j, 4) + gap <= p(i, 2) + 0.000001 Or _
                       p(i, 3) + p(i, 5) + gap <= p(j, 3) + 0.000001 Or _
                       p(j, 3) + p(j, 5) + gap <= p(i, 3) + 0.000001, name & " clearance"
            End If
        Next
    Next
    Assert parts = expParts, name & " physical parts count"
    Assert Abs(area - q.TotalArea) < 0.000001, name & " profile area conservation"

    ' Check layout determinism
    UINestQuote q, 0.25, 0.25
    Assert UBound(q.SheetPlacements) = UBound(p), name & " deterministic count"
    For i = 0 To UBound(p)
        For j = 0 To 8
            Assert p(i, j) = q.SheetPlacements(i, j), name & " deterministic placement field " & j
        Next
    Next
End Sub

Dim q
' Five reference examples reconstructed from attached PDFs:
' 1. rad290EE: 4 walls, 100x100x24, Single return 2 inches, Mild Steel, 1/4 inch
Set q = UICalculateQuote(4, "N/A", "Single return 2 inches", "Mild Steel", "1/4 inch", 100, 100, 24)
CheckExample q, "rad290EE (100x100x24 T.25)", 4, 4, 34, 124.98

' 2. rad34B17: 4 walls, 100x100x24, Single return 2 inches, Mild Steel, 3/16 inch
Set q = UICalculateQuote(4, "N/A", "Single return 2 inches", "Mild Steel", "3/16 inch", 100, 100, 24)
CheckExample q, "rad34B17 (100x100x24 T.1875)", 4, 4, 34, 125.49

' 3. radC91FD: 4 walls, 100x60x48, No return, Mild Steel, 1/4 inch
Set q = UICalculateQuote(4, "N/A", "No return", "Mild Steel", "1/4 inch", 100, 60, 48)
CheckExample q, "radC91FD (100x60x48 T.25)", 6, 6, 28, 222.12

' 4. rad9B80B: 4 walls, 500x650x12, Single return 1 inch, Mild Steel, 3/16 inch
Set q = UICalculateQuote(4, "N/A", "Single return 1 inch", "Mild Steel", "3/16 inch", 500, 650, 12)
CheckExample q, "rad9B80B (500x650x12 T.1875)", 8, 8, 180, 328.79

' 5. radC046F: 4 walls, 300x300x30, Double return, Mild Steel, 3/16 inch
Set q = UICalculateQuote(4, "N/A", "Double return", "Mild Steel", "3/16 inch", 300, 300, 30)
CheckExample q, "radC046F (300x300x30 T.1875)", 13, 13, 94, 501.03

WScript.Echo "PASS: five realistic reference examples and determinism"
"""
        with tempfile.TemporaryDirectory(prefix="test-vbs-") as folder:
            runner = Path(folder) / "run.vbs"
            runner.write_text(test_body, encoding="ascii")
            res = subprocess.run(["cscript", "//Nologo", str(runner)], capture_output=True, text=True, timeout=60)
            self.assertEqual(res.returncode, 0, res.stdout + res.stderr)
            self.assertIn("PASS: five realistic reference examples and determinism", res.stdout)
