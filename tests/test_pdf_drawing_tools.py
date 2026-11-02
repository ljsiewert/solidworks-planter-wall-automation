"""Drawing-layout math and export contracts; not a full VBA compilation test."""
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import unittest

from test_userform_sources import ROOT, helper_body

TOOLS = (ROOT / "src" / "PDFDrawingTools.bas").read_text(encoding="ascii")


class PDFContracts(unittest.TestCase):
    def test_balloon_clearance_runs_after_placement_and_final_rebuild(self):
        layout = helper_body("PDFLayoutViews", TOOLS)
        self.assertLess(layout.index("view.Position = newPosition"), layout.index("PDFClearAssemblyBalloons views, False"))
        self.assertLess(layout.index("PDFPositionExplodedBaseplates view1"), layout.index("PDFClearAssemblyBalloons views, False"))
        verify = helper_body("PDFVerifyViews", TOOLS)
        self.assertIn("PDFClearAssemblyBalloons views, True", verify)
        bounds = helper_body("PDFBalloonBounds", TOOLS)
        self.assertNotIn("GetBalloonInfo", bounds)
        self.assertIn("GetArcAtIndex2", bounds)
        arrange = helper_body("PDFClearAssemblyBalloons", TOOLS)
        self.assertIn("GetLeaderPointsAtIndex", arrange)
        moved = arrange[arrange.index("If annotation.SetPosition2"):]
        self.assertLess(moved.index("bounds = PDFBalloonBounds(annotation)"), moved.index("If bounds(0) < usable(0)"))
        self.assertNotIn("SetLeaderAttachment", arrange)
        self.assertNotIn("SetAttachedEntities", arrange)
        export = helper_body("PDFExportDrawing", TOOLS)
        self.assertLess(export.index("PDFClearPrintedBalloons drawing, False"), export.index(".Save3("))
        self.assertLess(export.index(".Save3("), export.index("PDFClearPrintedBalloons drawing, True"))
        self.assertLess(export.index("PDFClearPrintedBalloons drawing, True"), export.index(".Extension.SaveAs("))

    def test_layout_avoids_vba_reserved_scale_identifier(self):
        body = helper_body("PDFPlanViews", TOOLS)
        self.assertIn("scaleFactor", body)
        self.assertNotRegex(body, r"(?i)\bscale\b")

    def test_exploded_balloons_follow_bom_and_actual_components(self):
        body = helper_body("PDFRefreshExplodedBalloons", TOOLS)
        self.assertIn("swBalloonItemNumbers_DoNotChangeItemNumbers", body)
        self.assertIn("bom.GetComponentsCount2(", body)
        self.assertIn("bom.GetComponents2(", body)
        self.assertIn("entity.GetComponent", body)
        self.assertIn("expected.Exists(key)", body)
        self.assertIn("found.Exists(key)", body)
        self.assertIn("annotation.Visible = swAnnotationVisibilityState_e.swAnnotationVisible", body)
        self.assertNotIn('"6"', body)
        prepare = helper_body("PDFPrepareSheet", TOOLS)
        self.assertLess(prepare.index("PDFRefreshExplodedBalloons("), prepare.index("PDFLayoutViews "))
        position = helper_body("PDFPositionBaseplateBalloon", TOOLS)
        self.assertIn("component.GetTotalTransform(True)", position)
        self.assertIn("view.ModelToViewTransform", position)
        self.assertIn("GetVisibleEntities2", position)
        self.assertIn("SetLeaderAttachmentPointAtIndex", position)
        layout = helper_body("PDFLayoutViews", TOOLS)
        self.assertLess(layout.index("view.Position = newPosition"), layout.index("PDFPositionExplodedBaseplates view1"))
        self.assertIn('"\\1-S"', prepare)

    def test_visibility_filter_precedes_layout_and_selection(self):
        body = helper_body("PDFPrepareSheet", TOOLS)
        self.assertLess(body.index("PDFHidePrintArtifacts("), body.index("PDFLayoutViews "))
        self.assertIn("annotation.Visible <> swAnnotationVisibilityState_e.swAnnotationHidden", body)
        self.assertIn("; hidden dangling dimensions/center marks=", body)
        hide = helper_body("PDFHidePrintArtifacts", TOOLS)
        self.assertIn("annotation.IsDangling", hide)
        self.assertIn("annotation.GetType = swAnnotationType_e.swDisplayDimension", hide)
        self.assertIn("annotation.GetType = swAnnotationType_e.swCenterMarkSym", hide)
        self.assertNotIn(".Delete", hide)

    def test_checked_single_export_without_reference_saves(self):
        body = helper_body("PDFExportDrawing", TOOLS)
        self.assertEqual(body.count(".Extension.SaveAs("), 1)
        self.assertEqual(body.count(".Save3("), 1)
        self.assertIn("sheets = drawing.GetSheetNames", body)
        self.assertIn("swExportData_ExportSpecifiedSheets, sheets", body)
        self.assertIn("ViewPdfAfterSaving = False", body)
        self.assertIn("FileLen(path) = 0", body)
        self.assertNotIn("SaveReferenced", body)
        self.assertNotIn("On Error Resume Next", TOOLS)

    def test_both_exporters_use_measured_layout_and_single_export(self):
        for name, functions in (
            ("pdfassemblies", ("ProcessMainAssemblyDrawing", "ProcessSubAssemblyDrawing")),
            ("pdfcomponents", ("ProcessComponentDrawing",)),
        ):
            source = (ROOT / "src" / f"{name}.bas").read_text(encoding="utf-8")
            for function in functions:
                body = helper_body(function, source)
                self.assertIn("PDFPrepareSheet(", body)
                self.assertIn("PDFExportDrawing swApp, swDrawModel, pdfPath", body)
                self.assertNotIn("On Error Resume Next", body)
                self.assertNotIn(".SaveAs", body)
                self.assertNotIn("ExtractDimensions", body)
                self.assertNotIn("ForceRebuild3", body)

    def test_outline_anchor_and_annotation_guards(self):
        body = helper_body("PDFLayoutViews", TOOLS)
        self.assertIn("anchor(0) + centersX(i) - (bounds(0) + bounds(2)) / 2", body)
        self.assertIn("view.UseParentScale = False", body)
        self.assertIn("If scaleChanged Then", body)
        self.assertIn("annotation.GetPosition", helper_body("PDFViewBounds", TOOLS))
        self.assertIn("view.RemoveAlignment", body)
        self.assertIn("view.PositionLocked = False", body)
        self.assertIn("sheet.SetScale(", body)
        self.assertIn("View/annotation bounds overlap", TOOLS)
        prepare = helper_body("PDFPrepareSheet", TOOLS)
        self.assertEqual(prepare.count("ForceRebuild3"), 1)

    def test_component_counts_and_timings_are_real(self):
        source = (ROOT / "src" / "pdfcomponents.bas").read_text(encoding="utf-8")
        self.assertIn("successCount = successCount + 1", source)
        self.assertIn("failureCount = failureCount + 1", source)
        self.assertNotIn('fileCount & " component PDFs created successfully!"', source)
        for label in ("Open seconds:", "layout seconds:", "Save/export seconds:"):
            self.assertIn(label, source)


@unittest.skipUnless(shutil.which("cscript"), "Windows Script Host required")
class PDFLayoutMath(unittest.TestCase):
    def test_balloon_bounds_measure_printed_arcs_not_leaders_or_stale_note_coordinates(self):
        source = re.sub(r"As SldWorks\.\w+", "As Object", TOOLS)
        functions = "\n".join(
            helper_body(name, source) for name in ("PDFMin", "PDFMax", "PDFBalloonBounds")
        )
        script = r'''
Option Explicit
Const PDF_ERROR = 1001
Dim annotation, display, bounds
Class FakeDisplay
    Public Count
    Function GetArcCount()
        GetArcCount = Count
    End Function
    Function GetArcAtIndex2(i)
        GetArcAtIndex2 = Array(0,0,-1,-1, 2.1,3,0, 2.1,3,0, 2,3,0, 0,0,1,1)
    End Function
End Class
Class FakeAnnotation
    Public Display
    Function GetDisplayData()
        Set GetDisplayData = Display
    End Function
End Class
Set display = New FakeDisplay
display.Count = 1
Set annotation = New FakeAnnotation
Set annotation.Display = display
bounds = PDFBalloonBounds(annotation)
If Abs(bounds(0)-1.9)>0.000001 Or Abs(bounds(1)-2.9)>0.000001 Then Err.Raise 513, , "Lower circle bounds"
If Abs(bounds(2)-2.1)>0.000001 Or Abs(bounds(3)-3.1)>0.000001 Then Err.Raise 513, , "Upper circle bounds"
display.Count = 0
On Error Resume Next
bounds = PDFBalloonBounds(annotation)
If Err.Number <> PDF_ERROR Then WScript.Quit 1
On Error GoTo 0
'''
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "balloon-bounds.vbs"
            path.write_text(script + "\n" + functions, encoding="ascii")
            result = subprocess.run(
                ["cscript", "//nologo", str(path)], capture_output=True, text=True, timeout=30
            )
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_balloon_planner_keeps_circles_clear_of_all_geometry_and_each_other(self):
        functions = "\n".join(
            helper_body(name, TOOLS)
            for name in ("PDFMin", "PDFMax", "PDFBoxesOverlap", "PDFFindBalloonPosition")
        )
        functions = re.sub(r"(\d+)#", r"\1", functions)
        script = r'''
Option Explicit
Const PDF_ERROR = 1001
Const INCH = 0.0254
Dim geometry, occupied, size, own, ball, anchor, point, bounds, obstacle, i
Set occupied = CreateObject("Scripting.Dictionary")
For Each size In Array(0.5, 1, 2, 4)
    own = Array(0, 0, size, size / 2)
    geometry = Array(own, Array(size + 0.02, 0, size * 2, size / 2), Empty)
    occupied.RemoveAll
    ball = Array(-0.005, -0.005, 0.005, 0.005)
    For i = 0 To 11
        anchor = Array(size * (i Mod 4) / 3, size / 4)
        point = PDFFindBalloonPosition(own, geometry, occupied, ball, anchor, Empty)
        bounds = Array(point(0)-0.005, point(1)-0.005, point(0)+0.005, point(1)+0.005)
        For Each obstacle In geometry
            If IsArray(obstacle) Then
                If PDFBoxesOverlap(bounds, obstacle, 0.15 * INCH) Then Err.Raise 513, , "Circle overlaps model"
            End If
        Next
        For Each obstacle In occupied.Items
            If PDFBoxesOverlap(bounds, obstacle, 0.15 * INCH) Then Err.Raise 513, , "Circles overlap"
        Next
        occupied.Add occupied.Count, bounds
    Next
Next
occupied.RemoveAll
own = Array(0.1, 0.1, 0.3, 0.2)
geometry = Array(own, Empty, Empty)
point = PDFFindBalloonPosition(own, geometry, occupied, ball, Array(0.2,0.1), Array(0.08,0.095,0.35,0.25))
If point(1) < 0.095 + 0.005 Then Err.Raise 513, , "Balloon entered reserved footer"
If point(0) > 0.1 And point(0) < 0.3 And point(1) < 0.2 Then Err.Raise 513, , "Planner did not choose another side"
geometry = Array(Array(-10,-10,10,10), Empty, Empty)
On Error Resume Next
point = PDFFindBalloonPosition(own, geometry, occupied, ball, anchor, Empty)
If Err.Number <> PDF_ERROR Then WScript.Quit 1
On Error GoTo 0
'''
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "balloon-clearance.vbs"
            path.write_text(script + "\n" + functions, encoding="ascii")
            result = subprocess.run(
                ["cscript", "//nologo", str(path)], capture_output=True, text=True, timeout=30
            )
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_revision_suffix_does_not_turn_side_drawing_into_main_assembly(self):
        source = (ROOT / "src" / "pdfassemblies.bas").read_text(encoding="utf-8")
        body = helper_body("IsMainAssemblyDrawing", source)
        script = (
            body
            + '\n'
            + '\n'.join(
                f'If IsMainAssemblyDrawing("{name}") <> {str(expected)} Then Err.Raise 513, , "{name}"'
                for name, expected in (
                    ("1-S2400-2000-22-C1", False),
                    ("1-S2400-2000-22-C2", False),
                    ("1-C6975-2000-22-C1", False),
                    ("40-1-C15600-S2400-2000-22-C1", True),
                    ("40-1-C15600-S2400-2000-22-S1", True),
                    ("40-1-c15600-s2400-2000-22-c1", True),
                    ("invalid", False),
                )
            )
        )
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "classification.vbs"
            path.write_text(script, encoding="ascii")
            result = subprocess.run(
                ["cscript", "//nologo", str(path)], capture_output=True, text=True, timeout=30
            )
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_actual_balloon_helper_restores_gusset_and_rejects_missing_or_wrong_attachment(self):
        source = re.sub(r"As SldWorks\.\w+", "As Object", TOOLS)
        body = helper_body("PDFRefreshExplodedBalloons", source)
        for enum, value in (
            ("swTableAnnotationType_e.swTableAnnotation_BillOfMaterials", 2),
            ("swBalloonLayoutType_e.swDetailingBalloonLayout_Square", 1),
            ("swBalloonStyle_e.swBS_Circular", 1), ("swBalloonFit_e.swBF_2Chars", 2),
            ("swBalloonTextContent_e.swBalloonTextItemNumber", 1),
            ("swBalloonItemNumbersOrder_e.swBalloonItemNumbers_DoNotChangeItemNumbers", 1),
            ("swEditBalloonOption_e.swEditBalloonOption_Replace", 0),
            ("swAnnotationType_e.swNote", 6),
            ("swAnnotationVisibilityState_e.swAnnotationHidden", 3),
            ("swAnnotationVisibilityState_e.swAnnotationVisible", 1),
        ):
            body = body.replace(enum, str(value))
        body = body.replace("LCase$", "LCase").replace("0#", "0").replace("table.Type", "table.TableType")
        script = r'''
Option Explicit
Const PDF_ERROR = 1001
Dim withGusset, wrong, missing, drawing, model, view, first, second, stale, count
Class FakeComponent
    Public Path
    Function GetPathName()
        GetPathName = Path
    End Function
End Class
Class FakeEntity
    Public Component
    Function GetComponent()
        Set GetComponent = Component
    End Function
End Class
Class FakeNote
    Public Item
    Function IsBomBalloon()
        IsBomBalloon = True
    End Function
    Function GetBomBalloonText(upper)
        GetBomBalloonText = Item
    End Function
End Class
Class FakeAnnotation
    Public Visible, Dangling, Note, Entity, NextItem
    Function GetType()
        GetType = 6
    End Function
    Function GetSpecificAnnotation()
        Set GetSpecificAnnotation = Note
    End Function
    Function IsDangling()
        IsDangling = Dangling
    End Function
    Function GetAttachedEntities3()
        GetAttachedEntities3 = Array(Entity)
    End Function
    Function GetNext3()
        Set GetNext3 = NextItem
    End Function
End Class
Class FakeTable
    Public RowCount
    Public Property Get TableType()
        TableType = 2
    End Property
    Function GetComponentsCount2(row, config, item, part)
        GetComponentsCount2 = 0
        If row > 0 Then
            If row = 1 Then item = "1" Else item = "6"
            GetComponentsCount2 = 1
        End If
    End Function
    Function GetComponents2(row, config)
        Dim component
        Set component = New FakeComponent
        If row = 1 Then component.Path = "plate" Else component.Path = "gusset"
        GetComponents2 = Array(component)
    End Function
End Class
Class FakeView
    Public First
    Public Name, ReferencedConfiguration
    Function GetFirstAnnotation3()
        Set GetFirstAnnotation3 = First
    End Function
    Function GetTableAnnotations()
        Dim table
        Set table = New FakeTable
        table.RowCount = 2
        If withGusset Then table.RowCount = 3
        GetTableAnnotations = Array(table)
    End Function
End Class
Class FakeOptions
    Public Layout, IgnoreMultiple, LeaderAttachmentToFaces, Style, Size, UpperTextContent
    Public ItemOrder, InsertMagneticLine, EditBalloons, EditBalloonOption
End Class
Class FakeDrawing
    Public View
    Function GetFirstView()
        Set GetFirstView = View
    End Function
    Function CreateAutoBalloonOptions()
        Set CreateAutoBalloonOptions = New FakeOptions
    End Function
    Function AutoBalloon5(options)
        If options.ItemOrder <> 1 Then Err.Raise 1002, , "BOM was renumbered"
        AutoBalloon5 = Array(first)
    End Function
End Class
Class FakeModel
    Public Property Get Extension()
        Set Extension = Me
    End Property
    Function SelectByID2(name, kind, x, y, z, append, mark, callout, optionValue)
        SelectByID2 = 1
    End Function
    Sub ClearSelection2(all)
    End Sub
End Class
Function MakeAnnotation(item, path, dangling)
    Dim a, n, e, c
    Set a = New FakeAnnotation
    Set n = New FakeNote
    n.Item = item
    Set c = New FakeComponent
    c.Path = path
    Set e = New FakeEntity
    Set e.Component = c
    Set a.Note = n
    Set a.Entity = e
    Set a.NextItem = Nothing
    a.Visible = 3
    a.Dangling = dangling
    Set MakeAnnotation = a
End Function
Sub PDFPositionBaseplateBalloon(v, a, c)
    Err.Raise 1003, , "Unexpected baseplate geometry call for fake paths"
End Sub
Sub Check(gusset, bad, absent)
    withGusset = gusset
    Set first = MakeAnnotation("1", "plate", False)
    Set second = MakeAnnotation("6", "gusset", False)
    If bad Then second.Entity.Component.Path = "plate"
    Set stale = MakeAnnotation("?", "", True)
    stale.Visible = 1
    Set first.NextItem = stale
    If gusset And Not absent Then
        Set first.NextItem = second
        Set second.NextItem = stale
    End If
    Set view = New FakeView
    view.Name = "Exploded"
    view.ReferencedConfiguration = "Default"
    Set view.First = first
    Set drawing = New FakeDrawing
    Set drawing.View = view
    Set model = New FakeModel
    On Error Resume Next
    count = PDFRefreshExplodedBalloons(drawing, model, view)
    If bad Or absent Then
        If Err.Number = 0 Then WScript.Quit 1
        Err.Clear
    Else
        If Err.Number <> 0 Then
            WScript.Echo Err.Description
            WScript.Quit 1
        End If
        If first.Visible <> 1 Or stale.Visible <> 3 Then WScript.Quit 1
        If gusset Then
            If count <> 2 Or second.Visible <> 1 Then WScript.Quit 1
        Else
            If count <> 1 Then WScript.Quit 1
        End If
    End If
    On Error GoTo 0
End Sub
Check False, False, False
Check True, False, False
Check True, True, False
Check True, False, True
WScript.Echo "PASS: BOM-linked visibility, missing balloon and attachment guards"
''' + body
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "balloons.vbs"
            path.write_text(script, encoding="ascii")
            result = subprocess.run(["cscript", "//Nologo", str(path)], capture_output=True, text=True, timeout=30)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_only_dangling_dimensions_and_center_marks_are_hidden(self):
        source = re.sub(r"As SldWorks\.\w+", "As Object", TOOLS)
        helper = helper_body("PDFHidePrintArtifacts", source)
        helper = helper.replace("swAnnotationType_e.swDisplayDimension", "4")
        helper = helper.replace("swAnnotationType_e.swCenterMarkSym", "13")
        helper = helper.replace("swAnnotationVisibilityState_e.swAnnotationHidden", "3")
        helper = helper.replace("swUserPreferenceToggle_e.swDisplaySketches", "196")
        script = r'''
Option Explicit
Const PDF_ERROR = 1001
Dim items(5), i, view, model, drawing, hidden
Class FakeAnnotation
    Public Visible, Kind, Dangling, NextItem
    Function IsDangling()
        IsDangling = Dangling
    End Function
    Function GetType()
        GetType = Kind
    End Function
    Function GetNext3()
        Set GetNext3 = NextItem
    End Function
End Class
Class FakeView
    Public First
    Function GetFirstAnnotation3()
        Set GetFirstAnnotation3 = First
    End Function
    Function GetNextView()
        Set GetNextView = Nothing
    End Function
End Class
Class FakeDrawing
    Public First
    Function GetFirstView()
        Set GetFirstView = First
    End Function
End Class
Class FakeModel
    Public SketchVisible
    Function SetUserPreferenceToggle(flag, value)
        SketchVisible = value
        SetUserPreferenceToggle = 0
    End Function
    Function GetUserPreferenceToggle(flag)
        GetUserPreferenceToggle = SketchVisible
    End Function
End Class
For i = 0 To 5
    Set items(i) = New FakeAnnotation
    items(i).Visible = 1
    items(i).Kind = 4
    items(i).Dangling = False
    Set items(i).NextItem = Nothing
    If i > 0 Then Set items(i - 1).NextItem = items(i)
Next
items(1).Dangling = True
items(2).Visible = 3
items(3).Kind = 6
items(3).Dangling = True
items(4).Kind = 13
items(5).Kind = 13
items(5).Dangling = True
Set view = New FakeView
Set view.First = items(0)
Set drawing = New FakeDrawing
Set drawing.First = view
Set model = New FakeModel
model.SketchVisible = True
hidden = PDFHidePrintArtifacts(drawing, model)
If model.SketchVisible Or hidden <> 2 Or items(0).Visible <> 1 Or items(1).Visible <> 3 Or _
   items(2).Visible <> 3 Or items(3).Visible <> 1 Or items(4).Visible <> 1 Or items(5).Visible <> 3 Then
    WScript.Echo "FAIL: valid dimensions/center marks, hidden dimensions or notes altered"
    WScript.Quit 1
End If
If PDFHidePrintArtifacts(drawing, model) <> 0 Then
    WScript.Echo "FAIL: repeated cleanup double-counted hidden dimensions"
    WScript.Quit 1
End If
WScript.Echo "PASS: dangling-only filtering"
''' + helper
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "visibility.vbs"
            path.write_text(script, encoding="ascii")
            result = subprocess.run(["cscript", "//Nologo", str(path)], capture_output=True, text=True, timeout=30)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_dimension_lines_are_included_in_bounds(self):
        source = re.sub(r"As SldWorks\.\w+", "As Object", TOOLS)
        helpers = "\n\n".join(helper_body(name, source) for name in ("PDFMin", "PDFMax", "PDFViewBounds"))
        helpers = helpers.replace("swAnnotationType_e.swDisplayDimension", "4").replace("swAnnotationType_e.swNote", "6")
        helpers = helpers.replace("swAnnotationVisibilityState_e.swAnnotationHidden", "3")
        script = r'''
Option Explicit
Const PDF_ERROR = 1001
Const INCH = 0.0254
Dim view, bounds
Class FakeDisplay
    Function GetLineCount()
        GetLineCount = 1
    End Function
    Function GetLineAtIndex3(index)
        GetLineAtIndex3 = Array(65535,0,0,0,-1,0.75,0,4,3,0)
    End Function
End Class
Class FakeAnnotation
    Public Visible
    Private Sub Class_Initialize()
        Visible = 1
    End Sub
    Function GetPosition()
        GetPosition = Array(1.5,1.5,0)
    End Function
    Function GetType()
        GetType = 4
    End Function
    Function GetDisplayData()
        Set GetDisplayData = New FakeDisplay
    End Function
    Function GetNext3()
        Set GetNext3 = Nothing
    End Function
End Class
Class FakeView
    Public Hidden
    Function GetOutline()
        GetOutline = Array(1,1,2,2)
    End Function
    Function GetFirstAnnotation3()
        Dim item
        Set item = New FakeAnnotation
        If Hidden Then item.Visible = 3
        Set GetFirstAnnotation3 = item
    End Function
End Class
Set view = New FakeView
bounds = PDFViewBounds(view)
If bounds(0) <> -1 Or bounds(1) <> 0.75 Or bounds(2) <> 4 Or bounds(3) <> 3 Then
    WScript.Echo "FAIL: dimension line endpoints not included or style fields treated as coordinates"
    WScript.Quit 1
End If
view.Hidden = True
bounds = PDFViewBounds(view)
If bounds(0) <> 1 Or bounds(1) <> 1 Or bounds(2) <> 2 Or bounds(3) <> 2 Then
    WScript.Echo "FAIL: hidden annotations still shrink drawing views"
    WScript.Quit 1
End If
WScript.Echo "PASS: line endpoint bounds"
''' + helpers
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "bounds.vbs"
            path.write_text(script, encoding="ascii")
            result = subprocess.run(["cscript", "//Nologo", str(path)], capture_output=True, text=True, timeout=30)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_actual_plan_bounds_scales_and_nonoverlap(self):
        helpers = "\n\n".join(helper_body(name, TOOLS) for name in (
            "PDFStandardScale", "PDFMin", "PDFMax", "PDFPlanViews",
        ))
        helpers = re.sub(r"(\d)#", r"\1", helpers)
        script = r'''
Option Explicit
Const PDF_ERROR = 1001
Const INCH = 0.0254
Dim scales, cx, cy, w, h, kind, i, j, cases, c, failed
Sub Assert(condition, message)
    If Not condition Then
        WScript.Echo "FAIL: " & message
        WScript.Quit 1
    End If
End Sub
''' + helpers + r'''
cases = Array( _
    Array("component", Array(38.25, 0.25, 4), Array(20, 20, 0.5)), _
    Array("component", Array(7.5, 0.25, 0), Array(14.15, 14.15, 0)), _
    Array("component", Array(69.75, 0.25, 0), Array(20, 20, 0)), _
    Array("component", Array(8, 0, 0), Array(69.75, 0, 0)), _
    Array("component", Array(5, 0.1, 50), Array(8, 8, 1)), _
    Array("main", Array(156, 156, 148), Array(24, 20, 100)), _
    Array("exploded", Array(100, 0, 0), Array(60, 0, 0)), _
    Array("assembly", Array(60, 70, 0), Array(40, 8, 0)), _
    Array("assembly", Array(0, 24, 0), Array(0, 5, 0)))
For Each c In cases
    kind = c(0): w = c(1): h = c(2)
    For i = 0 To 2
        w(i) = w(i) * INCH
        h(i) = h(i) * INCH
    Next
    PDFPlanViews kind, w, h, 17 * INCH, 11 * INCH, scales, cx, cy
    For i = 0 To 2
        If w(i) > 0 Then
            Assert scales(i) > 0 And scales(i) <= 1, "Valid scale: " & kind
            Assert cx(i) - w(i) * scales(i) / 2 >= 0.85 * INCH - 0.000001, "Left boundary: " & kind
            Assert cx(i) + w(i) * scales(i) / 2 <= 14.5 * INCH + 0.000001, "Title block boundary: " & kind
            Assert cy(i) - h(i) * scales(i) / 2 >= 0.85 * INCH - 0.000001, "Bottom boundary: " & kind
            Assert cy(i) + h(i) * scales(i) / 2 <= 10.15 * INCH + 0.000001, "Top boundary: " & kind
            If kind = "main" Or kind = "exploded" Then
                Assert cy(i) + h(i) * scales(i) / 2 <= 8.55 * INCH + 0.000001, "BOM clearance"
            End If
            For j = i + 1 To 2
                If w(j) > 0 Then
                    Assert Abs(cx(i) - cx(j)) >= (w(i) * scales(i) + w(j) * scales(j)) / 2 Or _
                        Abs(cy(i) - cy(j)) >= (h(i) * scales(i) + h(j) * scales(j)) / 2, "Nonoverlap: " & kind
                End If
            Next
        End If
    Next
    If kind = "component" Then
        For i = 1 To 2
            If w(i) > 0 Then Assert scales(i) = scales(0), "Common component scale"
        Next
    End If
    If kind = "main" Or kind = "assembly" Then
        For i = 0 To 2
            For j = i + 1 To 2
                If w(i) > 0 And w(j) > 0 Then Assert scales(i) = scales(j), "Common sheet scale"
            Next
        Next
    End If
Next
Assert PDFStandardScale(0.13) = 0.125, "Standard scale rounds down"
On Error Resume Next
PDFPlanViews "component", Array(0,0,0), Array(0,0,0), 0.4, 0.3, scales, cx, cy
failed = Err.Number
On Error GoTo 0
Assert failed <> 0, "Empty layout rejected"
WScript.Echo "PASS: sheet layout planning"
'''
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "layout.vbs"
            path.write_text(script, encoding="ascii")
            result = subprocess.run(["cscript", "//Nologo", str(path)], capture_output=True, text=True, timeout=30)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)


if __name__ == "__main__":
    unittest.main()
