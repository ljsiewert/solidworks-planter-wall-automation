"""Field editing, common-margin behavior and guarded selected-preview contracts."""
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
FORM = (ROOT / "src" / "UserForm_AutomationUI.frm").read_text(encoding="ascii")
FIELD_FORM = (ROOT / "src" / "UserForm_FieldQuote.frm").read_text(encoding="ascii")
FIELD = (ROOT / "src" / "ui_quote_field.bas").read_text(encoding="ascii")
MODULE = (ROOT / "src" / "userform_automation.bas").read_text(encoding="ascii")
CALCULATOR = (ROOT / "src" / "ui_quote_calculator.bas").read_text(encoding="ascii")


def procedure(source, name):
    body = re.search(rf"^(?:Public|Private) (Sub|Function) {name}\b.*?^End \1$",
                     source, re.MULTILINE | re.DOTALL)[0]
    body = re.sub(r" As (?:Long|Double|Boolean|String|UIQuoteResult|Collection)\b", "", body)
    body = re.sub(r"^(Public|Private) ", "", body, flags=re.MULTILINE)
    body = re.sub(r"^(\s*Next) \w+$", r"\1", body, flags=re.MULTILINE)
    return body


class FieldEditorTests(unittest.TestCase):
    def test_quote_run_requires_saved_designs_and_skips_draft(self):
        run = procedure(FORM, "runButton_Click")
        quote_path = run.split("If quoteBox.Value Then", 1)[1].split(
            "    If wallsBox.ListIndex < 0 Then", 1)[0]
        self.assertLess(quote_path.index("If fieldDesigns.Count = 0 Then"),
                        quote_path.index("Set quoteDesigns = fieldDesigns"))
        self.assertIn("Err.Raise", quote_path)
        self.assertIn("UIFieldSetMargin fieldDesigns, quoteMargin", quote_path)
        self.assertIn("Me.Hide", quote_path)
        self.assertIn("Exit Sub", quote_path)
        self.assertNotIn("ReadQuoteDesign", run)
        self.assertNotIn("UIShowQuote quoteValues", MODULE)
        self.assertIn('AddCheck("chkQuoteOnly", "Quote only", 16, 344)', FORM)
        self.assertIn('AddCheck("chkPreviewOnly", "Preview only (no files created)", 235, 344)', FORM)
        self.assertNotIn('AddText("txtQuantity"', FORM)

    def test_edit_is_validated_before_replacing_saved_snapshot(self):
        save = procedure(FORM, "updatePlanterButton_Click")
        self.assertLess(save.index("Set result = ReadQuoteDesign()"),
                        save.index("UIFieldReplaceDesign fieldDesigns"))
        self.assertIn("RefreshDesignList selected", save)
        self.assertIn('MsgBox Err.Description, vbExclamation, "Review Planter Design"', save)
        read = procedure(FORM, "ReadQuoteDesign")
        for expected in ("UIValidateSelections", "UIValidateDimensions",
                         "UIQuoteParseQuantity(quantityBox.Text)", "UICalculateQuote",
                         "UIQuoteSetQuantity result, quantity",
                         "result.EstimateLabel = Trim$(designNameBox.Text)"):
            self.assertIn(expected, read)
        self.assertLess(read.index("UIQuoteSetQuantity result, quantity"),
                        read.index("Set ReadQuoteDesign = result"))
        self.assertNotIn("fieldDesigns", read)

    def test_pricing_cards_do_not_overlap_or_overrun_table(self):
        cards = re.findall(
            r'AddPriceCard\("\w+", "[^"]+", (\d+), (\d+), (\d+), (\d+), (\d+), (True|False)\)',
            FIELD_FORM)
        self.assertEqual(len(cards), 7)
        boxes = []
        for left, top, width, height, font, _ in cards:
            left, top, width, height, font = map(int, (left, top, width, height, font))
            self.assertLessEqual(left + width, 762)
            self.assertLessEqual(top + height, 242)
            self.assertGreaterEqual(height - 25, font + 2)
            boxes.append((left, top, left + width, top + height))
        for i, (left, top, right, bottom) in enumerate(boxes):
            for other_left, other_top, other_right, other_bottom in boxes[i + 1:]:
                self.assertTrue(right <= other_left or other_right <= left or
                                bottom <= other_top or other_bottom <= top)

    def test_selected_preview_reuses_guards_and_restores_field_context(self):
        preview = procedure(MODULE, "UIPreviewQuoteDesign")
        self.assertIn("packngo1.Step1_Initialize()", preview)
        self.assertNotIn("UIRequireCleanTemplate", preview)
        self.assertIn('"planter_assembly.SLDASM"', preview)
        self.assertLess(preview.index('"planter_assembly.SLDASM"'), preview.index("UIMapSelections"))
        self.assertIn("UIRunPreview", preview)
        for forbidden in ("Step8", "UIRunAutomation", "Save3", "OpenDoc6", "UICreateDestination"):
            self.assertNotIn(forbidden, preview)
        for value in ("OverallLength", "OverallWidth", "Height", "WallCount",
                      "Orientation", "ReturnLabel", "Material", "Thickness"):
            self.assertIn("result." + value, preview)
        self.assertNotIn("result.Quantity", preview)
        for state, old in (("designTableInputs", "previousInputs"), ("quoteOnly", "previousQuoteOnly"),
                           ("previewOnly", "previousPreviewOnly"), ("outputFolder", "previousOutput")):
            self.assertIn(f"{state} = {old}", preview)
        self.assertIn("Resume Finished", preview)
        show = procedure(FIELD, "UIShowFieldQuote")
        self.assertEqual(show.count("Set report = New UserForm_FieldQuote"), 1)
        self.assertLess(show.index("Set report = New UserForm_FieldQuote"),
                        show.index("Do\n"))
        self.assertLess(show.index("report.SetDesigns designs"), show.index("Do\n"))
        self.assertIn("previewIndex = report.PreviewDesignIndex", show)
        self.assertIn("UIPreviewQuoteDesign designs(previewIndex)", show)
        self.assertGreater(show.index("Unload report"), show.index("Loop"))
        self.assertGreater(show.index("Set report = Nothing"), show.index("Loop"))
        self.assertIn("Do\n", show)
        self.assertIn("Loop", show)
        request = procedure(FIELD_FORM, "previewButton_Click")
        self.assertIn("mPreviewDesignIndex = designList.ListIndex + 1", request)
        self.assertIn("Me.Hide", request)
        self.assertIn("Public Property Get PreviewDesignIndex() As Long", FIELD_FORM)
        self.assertNotIn("PreviewDesign() As UIQuoteResult", FIELD_FORM)
        self.assertIn("Error \" & CStr(Err.Number) & \" in \" & Err.Source", MODULE)

    def test_field_margin_is_live_and_detail_margin_is_locked(self):
        self.assertIn("Private WithEvents marginBox As MSForms.TextBox", FIELD_FORM)
        self.assertIn("Private WithEvents marginSpin As MSForms.SpinButton", FIELD_FORM)
        self.assertIn("marginSpin.Min = 0", FIELD_FORM)
        self.assertIn("marginSpin.Max = 95", FIELD_FORM)
        for name in ("marginBox_Change", "marginSpin_Change"):
            self.assertIn("ApplyMargin", procedure(FIELD_FORM, name))
        self.assertIn("UIShowQuote result, False", FIELD_FORM)
        self.assertIn('Set metricValues(0) = AddPriceCard("Sell", "FIELD SELL PRICE", 12, 42, 750, 76, 28, True)', FIELD_FORM)
        self.assertIn('AddPriceCard("Profit", "PROFIT", 12, 126, 371, 60, 20, False)', FIELD_FORM)
        self.assertIn('AddPriceCard("Material", "MATERIAL", 12, 194, 182, 48, 14, False)', FIELD_FORM)
        self.assertIn('indices = Array(7, 6, 5, 2, 3, 4, 1)', FIELD_FORM)
        self.assertIn('metricValues(i).Caption = "Withheld"', FIELD_FORM)
        self.assertNotIn("marginButton", FIELD_FORM)

    @unittest.skipUnless(shutil.which("cscript"), "Windows Script Host required")
    def test_loading_selection_replacement_and_shared_margin_math(self):
        result = (ROOT / "src" / "UIQuoteResult.cls").read_text(encoding="ascii")
        fields = re.sub(r" As (?:Long|String|Double|Boolean|Variant)\b", "",
                        result.split("Option Explicit", 1)[1])
        script = r'''
Option Explicit
Dim loadingDesign, fieldDesigns, planterList, updatePlanterButton, quoteBox, designNameBox, quantityBox
Dim wallsBox, sideBox, returnBox, lengthBox, widthBox, heightBox, materialBox, thicknessBox, paintedToggle
Dim priceBox, cutRateBox, bendRateBox, weldRateBox, paintRateBox, gapBox, marginBox, quoteMarginBox, nextDesignNumber
Dim mountBox, overrideBox, baseWidthBox, customBox
Dim a, b, replacement, code, design
Class Control
    Public Text, Value, Enabled, ListIndex
End Class
Class DesignList
    Private entries(2)
    Public Count
    Public Sub Init(first, second)
        Set entries(0) = first
        Set entries(1) = second
        Count = 2
    End Sub
    Public Default Property Get Item(index)
        Set Item = entries(index - 1)
    End Property
    Public Sub Insert(value, index)
        Dim i
        For i = Count To index Step -1
            Set entries(i) = entries(i - 1)
        Next
        Set entries(index - 1) = value
        Count = Count + 1
    End Sub
    Public Sub Remove(index)
        Dim i
        For i = index - 1 To Count - 2
            Set entries(i) = entries(i + 1)
        Next
        Count = Count - 1
    End Sub
    Public Function Items()
        Items = Array(entries(0), entries(1))
    End Function
End Class
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
        script += procedure(FORM, "planterList_Change") + "\n"
        script += procedure(FORM, "newPlanterButton_Click") + "\n"
        script += procedure(FIELD, "UIFieldReplaceDesign").replace(
            "designs.Add result, Before:=index", "designs.Insert result, index") + "\n"
        script += procedure(FIELD, "UIFieldSetMargin").replace(
            "For Each result In designs", "For Each result In designs.Items") + "\n"
        script += procedure(CALCULATOR, "UIQuoteSetMargin") + "\n"
        for name in ("planterList", "updatePlanterButton", "quoteBox", "designNameBox", "quantityBox",
                     "wallsBox", "sideBox", "returnBox", "lengthBox", "widthBox", "heightBox",
                     "materialBox", "thicknessBox", "paintedToggle", "priceBox", "cutRateBox",
                     "bendRateBox", "weldRateBox", "paintRateBox", "gapBox", "marginBox", "quoteMarginBox",
                     "mountBox", "overrideBox", "baseWidthBox", "customBox"):
            script += f"Set {name} = New Control\n"
        script += r'''
Set a = New UIQuoteResult
Set b = New UIQuoteResult
a.EstimateLabel = "North"
a.Quantity = 3
a.WallCount = 2
a.Orientation = "Left"
a.ReturnLabel = "Double return"
a.OverallLength = 350.125
a.OverallWidth = 100.25
a.Height = 20
a.Material = "Borcon Weathering Steel"
a.Thickness = 0.1875
a.Painted = True
a.PricePerPound = 0.95
a.CutRate = 80
a.BendRate = 6
a.WeldRate = 12
a.PaintRate = 3
a.PartGap = 0
a.EdgeMargin = 0.2
a.TotalCost = 1000
a.MountsToPlanter = True
a.CustomCuts = True
a.OverrideBaseWidth = True
a.BaseWidth = 7.5
b.EstimateLabel = "South"
b.TotalCost = 500
Set fieldDesigns = New DesignList
fieldDesigns.Init a, b
quoteBox.Value = True
quoteMarginBox.Text = "50"
planterList.ListIndex = 0
planterList_Change
Assert designNameBox.Text = "North" And quantityBox.Text = "3", "Selection loads saved name and quantity"
Assert wallsBox.Value = "2" And sideBox.Value = "Left" And returnBox.Value = a.ReturnLabel, "Selection loads design choices"
Assert CDbl(lengthBox.Text) = a.OverallLength And CDbl(widthBox.Text) = a.OverallWidth And CDbl(heightBox.Text) = a.Height, "Selection preserves dimensions"
Assert materialBox.Value = a.Material And thicknessBox.Value = "3/16 inch" And paintedToggle.Value, "Selection loads material and paint"
Assert CDbl(priceBox.Text) = a.PricePerPound And CDbl(cutRateBox.Text) = 80 And CDbl(bendRateBox.Text) = 6, "Selection loads prices"
Assert CDbl(weldRateBox.Text) = 12 And CDbl(paintRateBox.Text) = 3 And CDbl(gapBox.Text) = 0 And CDbl(marginBox.Text) = 0.2, "Selection loads rates and spacing"
Assert mountBox.Value And overrideBox.Value And CDbl(baseWidthBox.Text) = 7.5 And customBox.Value, "Selection loads mounting and baseplate options"
Assert quoteMarginBox.Text = "50", "Selecting designs never replaces shared margin"
Assert updatePlanterButton.Enabled, "Selection enables Save changes"
Set replacement = New UIQuoteResult
replacement.EstimateLabel = "North revised"
replacement.Quantity = 4
replacement.TotalCost = 1600
UIFieldReplaceDesign fieldDesigns, 1, replacement
Assert fieldDesigns.Count = 2, "Replacing a row preserves design count"
Assert fieldDesigns(1) Is replacement, "Replacement stays at selected index"
Assert fieldDesigns(2) Is b, "Other designs keep their identity and order"
Assert a.Quantity = 3, "Original snapshot is not mutated"
planterList_Change
Assert designNameBox.Text = "North revised" And quantityBox.Text = "4", "Edited name and quantity reload"
UIFieldSetMargin fieldDesigns, 50
Assert replacement.SellPrice = 3200 And b.SellPrice = 1000, "Shared 50% margin prices every design"
UIFieldSetMargin fieldDesigns, 40
Assert Abs(replacement.SellPrice - 1600 / 0.6) < 0.000001 And Abs(b.SellPrice - 500 / 0.6) < 0.000001, "Shared live margin math"
On Error Resume Next
UIFieldSetMargin fieldDesigns, 96
code = Err.Number
Err.Clear
On Error GoTo 0
Assert code <> 0 And replacement.MarginPercent = 40 And b.MarginPercent = 40, "Invalid field margin leaves all designs unchanged"
On Error Resume Next
UIFieldReplaceDesign fieldDesigns, 0, a
code = Err.Number
Err.Clear
On Error GoTo 0
Assert code <> 0 And fieldDesigns.Count = 2, "Invalid replacement index preserves collection"
nextDesignNumber = 5
newPlanterButton_Click
Assert planterList.ListIndex = -1 And Not updatePlanterButton.Enabled, "New design deselects saved row"
Assert designNameBox.Text = "Planter design 5" And quantityBox.Text = "1", "New design resets identity and quantity"
Assert fieldDesigns.Count = 2, "Draft does not add a saved design"
WScript.Echo "PASS: field editing and shared margins"
'''
        with tempfile.TemporaryDirectory(prefix="planter-field-") as folder:
            runner = Path(folder) / "field.vbs"
            runner.write_text(script, encoding="ascii")
            completed = subprocess.run(["cscript", "//Nologo", str(runner)],
                                       capture_output=True, text=True, timeout=30)
            self.assertEqual(completed.returncode, 0, completed.stdout + completed.stderr)
            self.assertIn("PASS:", completed.stdout)
