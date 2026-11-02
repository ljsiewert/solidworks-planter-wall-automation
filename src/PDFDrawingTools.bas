Attribute VB_Name = "PDFDrawingTools"
Option Explicit

' Import into both PDF exporter projects.
Private Const PDF_ERROR As Long = vbObjectError + 2600
Private Const INCH As Double = 0.0254

Public Function PDFElapsedSeconds(ByVal started As Double) As Double
    PDFElapsedSeconds = Timer - started
    If PDFElapsedSeconds < 0 Then PDFElapsedSeconds = PDFElapsedSeconds + 86400#
End Function

Public Function PDFStandardScale(ByVal maximum As Double) As Double
    Dim denominator As Variant
    If maximum <= 0 Then Err.Raise PDF_ERROR, "PDFStandardScale", "No usable drawing area."
    For Each denominator In Array(1, 2, 4, 5, 6, 8, 10, 12, 16, 20, 24, 30, 40, 50, 60, 80, 100, 150, 200, 300, 400, 500, 1000)
        If 1# / CDbl(denominator) <= maximum + 0.000000001 Then
            PDFStandardScale = 1# / CDbl(denominator)
            Exit Function
        End If
    Next denominator
    Err.Raise PDF_ERROR, "PDFStandardScale", "The model is too large for the drawing sheet."
End Function

Public Function PDFMin(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then PDFMin = a Else PDFMin = b
End Function

Public Function PDFMax(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then PDFMax = a Else PDFMax = b
End Function

Public Sub PDFPlanViews(ByVal kind As String, ByVal widths As Variant, ByVal heights As Variant, _
                        ByVal sheetWidth As Double, ByVal sheetHeight As Double, _
                        ByRef scales As Variant, ByRef centersX As Variant, ByRef centersY As Variant)
    Dim left As Double, bottom As Double, right As Double, top As Double
    Dim usableWidth As Double, usableHeight As Double, gap As Double
    Dim rowWidth As Double, rowHeight As Double, groupHeight As Double
    Dim scaleFactor As Double, startX As Double, startY As Double, rowBottom As Double
    Dim cellLeft As Variant, cellBottom As Variant, cellWidth As Variant, cellHeight As Variant
    Dim i As Long
    left = 0.85 * INCH
    bottom = 0.85 * INCH
    right = sheetWidth - 2.5 * INCH
    top = sheetHeight - 0.85 * INCH
    If kind = "main" Or kind = "exploded" Then top = top - 1.6 * INCH
    usableWidth = right - left
    usableHeight = top - bottom
    gap = 0.8 * INCH
    If usableWidth <= gap Or usableHeight <= gap Then
        Err.Raise PDF_ERROR, "PDFPlanViews", "Drawing sheet is too small for the reserved border/title-block areas."
    End If
    scales = Array(0#, 0#, 0#)
    centersX = Array(0#, 0#, 0#)
    centersY = Array(0#, 0#, 0#)

    If kind = "component" Then
        rowWidth = widths(0) + widths(1)
        rowHeight = PDFMax(heights(0), heights(1))
        groupHeight = rowHeight + heights(2)
        If rowWidth <= 0 Or groupHeight <= 0 Then Err.Raise PDF_ERROR, "PDFPlanViews", "No measurable component views."
        scaleFactor = PDFMin((usableWidth - gap) / PDFMax(rowWidth, widths(2)), (usableHeight - gap) / groupHeight)
        scaleFactor = PDFStandardScale(scaleFactor)
        rowWidth = rowWidth * scaleFactor
        If widths(1) > 0 Then rowWidth = rowWidth + gap
        groupHeight = groupHeight * scaleFactor
        If widths(2) > 0 Then groupHeight = groupHeight + gap
        startX = left + (usableWidth - PDFMax(rowWidth, widths(2) * scaleFactor)) / 2#
        startY = bottom + (usableHeight - groupHeight) / 2#
        rowBottom = startY
        If widths(2) > 0 Then rowBottom = rowBottom + heights(2) * scaleFactor + gap
        centersX(0) = startX + widths(0) * scaleFactor / 2#
        centersY(0) = rowBottom + rowHeight * scaleFactor / 2#
        centersX(1) = startX + widths(0) * scaleFactor + gap + widths(1) * scaleFactor / 2#
        centersY(1) = centersY(0)
        centersX(2) = startX + widths(2) * scaleFactor / 2#
        centersY(2) = startY + heights(2) * scaleFactor / 2#
        For i = 0 To 2
            If widths(i) > 0 Then scales(i) = scaleFactor
        Next i
        Exit Sub
    End If

    cellLeft = Array(left, left, left)
    cellBottom = Array(bottom, bottom, bottom)
    cellWidth = Array(usableWidth, usableWidth, usableWidth)
    cellHeight = Array(usableHeight, usableHeight, usableHeight)
    Select Case kind
        Case "main"
            cellWidth(0) = (usableWidth - gap) * 0.54
            cellWidth(1) = cellWidth(0)
            cellWidth(2) = usableWidth - gap - cellWidth(0)
            cellLeft(2) = left + cellWidth(0) + gap
            cellHeight(0) = (usableHeight - gap) * 0.6
            cellHeight(1) = usableHeight - gap - cellHeight(0)
            cellBottom(0) = bottom + cellHeight(1) + gap
        Case "assembly"
            If widths(0) > 0 And widths(1) > 0 Then
                cellHeight(0) = (usableHeight - gap) * 0.6
                cellHeight(1) = usableHeight - gap - cellHeight(0)
                cellBottom(0) = bottom + cellHeight(1) + gap
            End If
        Case "exploded"
        Case Else
            Err.Raise PDF_ERROR, "PDFPlanViews", "Unknown layout: " & kind
    End Select
    For i = 0 To 2
        If widths(i) > 0 And heights(i) > 0 Then
            scales(i) = PDFStandardScale(PDFMin(cellWidth(i) / widths(i), cellHeight(i) / heights(i)))
            centersX(i) = cellLeft(i) + cellWidth(i) / 2#
            centersY(i) = cellBottom(i) + cellHeight(i) / 2#
        End If
    Next i
    If kind = "main" Or kind = "assembly" Then
        scaleFactor = 1#
        For i = 0 To 2
            If scales(i) > 0 Then scaleFactor = PDFMin(scaleFactor, scales(i))
        Next i
        For i = 0 To 2
            If scales(i) > 0 Then scales(i) = scaleFactor
        Next i
    End If
End Sub

Public Function PDFViewBounds(ByVal view As SldWorks.View) As Variant
    Dim bounds As Variant, position As Variant
    Dim annotation As SldWorks.Annotation
    Dim note As SldWorks.Note, balloonBounds As Variant
    Dim display As SldWorks.DisplayData
    Dim line As Variant
    Dim i As Long, firstPoint As Long
    bounds = view.GetOutline
    If Not IsArray(bounds) Then Err.Raise PDF_ERROR, "PDFViewBounds", "No outline for view " & view.Name
    Set annotation = view.GetFirstAnnotation3
    Do While Not annotation Is Nothing
        If annotation.Visible <> swAnnotationVisibilityState_e.swAnnotationHidden Then
            If annotation.GetType = swAnnotationType_e.swNote Then
                Set note = annotation.GetSpecificAnnotation
                If note.IsBomBalloon Then
                    balloonBounds = PDFBalloonBounds(annotation)
                    bounds(0) = PDFMin(bounds(0), balloonBounds(0))
                    bounds(1) = PDFMin(bounds(1), balloonBounds(1))
                    bounds(2) = PDFMax(bounds(2), balloonBounds(2))
                    bounds(3) = PDFMax(bounds(3), balloonBounds(3))
                End If
            End If
            position = annotation.GetPosition
            If IsArray(position) Then
                bounds(0) = PDFMin(bounds(0), CDbl(position(0)) - 0.25 * INCH)
                bounds(1) = PDFMin(bounds(1), CDbl(position(1)) - 0.15 * INCH)
                bounds(2) = PDFMax(bounds(2), CDbl(position(0)) + 0.25 * INCH)
                bounds(3) = PDFMax(bounds(3), CDbl(position(1)) + 0.15 * INCH)
            End If
            Set display = Nothing
            If annotation.GetType = swAnnotationType_e.swDisplayDimension Or annotation.GetType = swAnnotationType_e.swNote Then
                Set display = annotation.GetDisplayData
            End If
            If Not display Is Nothing Then
                For i = 0 To display.GetLineCount - 1
                    line = display.GetLineAtIndex3(i)
                    If IsArray(line) Then
                        ' Display lines end with the start/end XYZ points after four style fields.
                        firstPoint = UBound(line) - 5
                        bounds(0) = PDFMin(bounds(0), PDFMin(CDbl(line(firstPoint)), CDbl(line(firstPoint + 3))))
                        bounds(1) = PDFMin(bounds(1), PDFMin(CDbl(line(firstPoint + 1)), CDbl(line(firstPoint + 4))))
                        bounds(2) = PDFMax(bounds(2), PDFMax(CDbl(line(firstPoint)), CDbl(line(firstPoint + 3))))
                        bounds(3) = PDFMax(bounds(3), PDFMax(CDbl(line(firstPoint + 1)), CDbl(line(firstPoint + 4))))
                    End If
                Next i
            End If
        End If
        Set annotation = annotation.GetNext3
    Loop
    PDFViewBounds = bounds
End Function

Public Function PDFBalloonBounds(ByVal annotation As SldWorks.Annotation) As Variant
    Dim display As SldWorks.DisplayData
    Dim arc As Variant, bounds As Variant
    Dim radius As Double, i As Long
    Set display = annotation.GetDisplayData
    If display Is Nothing Then Err.Raise PDF_ERROR, "PDFBalloonBounds", "Cannot read printed balloon geometry."
    For i = 0 To display.GetArcCount - 1
        arc = display.GetArcAtIndex2(i)
        If Not IsArray(arc) Then Err.Raise PDF_ERROR, "PDFBalloonBounds", "Cannot read balloon circle."
        ' Four style fields precede start/end XYZ, center XYZ and normal XYZ.
        radius = Sqr((arc(4) - arc(10)) ^ 2 + (arc(5) - arc(11)) ^ 2 + (arc(6) - arc(12)) ^ 2)
        If radius <= 0 Then Err.Raise PDF_ERROR, "PDFBalloonBounds", "Invalid printed balloon radius."
        If IsEmpty(bounds) Then
            bounds = Array(arc(10) - radius, arc(11) - radius, arc(10) + radius, arc(11) + radius)
        Else
            bounds(0) = PDFMin(bounds(0), arc(10) - radius)
            bounds(1) = PDFMin(bounds(1), arc(11) - radius)
            bounds(2) = PDFMax(bounds(2), arc(10) + radius)
            bounds(3) = PDFMax(bounds(3), arc(11) + radius)
        End If
    Next i
    If IsEmpty(bounds) Then Err.Raise PDF_ERROR, "PDFBalloonBounds", "BOM balloon has no measurable circle; use circular BOM balloons."
    PDFBalloonBounds = bounds
End Function

Public Function PDFBoxesOverlap(ByVal first As Variant, ByVal second As Variant, ByVal clearance As Double) As Boolean
    PDFBoxesOverlap = (first(0) < second(2) + clearance - 0.000001 And _
        first(2) > second(0) - clearance + 0.000001 And _
        first(1) < second(3) + clearance - 0.000001 And _
        first(3) > second(1) - clearance + 0.000001)
End Function

Public Function PDFFindBalloonPosition(ByVal ownBounds As Variant, ByVal geometry As Variant, _
                                       ByVal occupied As Object, ByVal balloonBounds As Variant, _
                                       ByVal anchor As Variant, ByVal usable As Variant) As Variant
    Dim halfWidth As Double, halfHeight As Double, gap As Double
    Dim x As Double, y As Double, fraction As Double, distance As Double, bestDistance As Double
    Dim sideIndex As Long, slot As Long, blocked As Boolean
    Dim candidate As Variant, obstacle As Variant, best As Variant
    gap = 0.2 * INCH
    halfWidth = (balloonBounds(2) - balloonBounds(0)) / 2#
    halfHeight = (balloonBounds(3) - balloonBounds(1)) / 2#
    For sideIndex = 0 To 3
        For slot = -1 To 32
            fraction = PDFMax(0#, slot) / 32#
            If sideIndex < 2 Then
                y = ownBounds(1) + fraction * (ownBounds(3) - ownBounds(1))
                If slot = -1 Then y = PDFMax(ownBounds(1), PDFMin(ownBounds(3), anchor(1)))
                If sideIndex = 0 Then
                    x = ownBounds(0) - gap - halfWidth
                Else
                    x = ownBounds(2) + gap + halfWidth
                End If
            Else
                x = ownBounds(0) + fraction * (ownBounds(2) - ownBounds(0))
                If slot = -1 Then x = PDFMax(ownBounds(0), PDFMin(ownBounds(2), anchor(0)))
                If sideIndex = 2 Then
                    y = ownBounds(1) - gap - halfHeight
                Else
                    y = ownBounds(3) + gap + halfHeight
                End If
            End If
            candidate = Array(x - halfWidth, y - halfHeight, x + halfWidth, y + halfHeight)
            blocked = False
            If IsArray(usable) Then
                If candidate(0) < usable(0) + 0.025 * INCH Or candidate(1) < usable(1) + 0.025 * INCH Or _
                   candidate(2) > usable(2) - 0.025 * INCH Or candidate(3) > usable(3) - 0.025 * INCH Then blocked = True
            End If
            For Each obstacle In geometry
                If IsArray(obstacle) Then
                    If PDFBoxesOverlap(candidate, obstacle, gap) Then blocked = True
                End If
            Next obstacle
            For Each obstacle In occupied.Items
                If PDFBoxesOverlap(candidate, obstacle, gap) Then blocked = True
            Next obstacle
            If Not blocked Then
                distance = (x - anchor(0)) ^ 2 + (y - anchor(1)) ^ 2
                If IsEmpty(best) Then
                    best = Array(x, y)
                    bestDistance = distance
                ElseIf distance < bestDistance Then
                    best = Array(x, y)
                    bestDistance = distance
                End If
            End If
        Next slot
    Next sideIndex
    If IsEmpty(best) Then Err.Raise PDF_ERROR, "PDFFindBalloonPosition", "No clear position outside the assembly for a BOM balloon."
    PDFFindBalloonPosition = best
End Function

Public Sub PDFClearAssemblyBalloons(ByVal views As Variant, ByVal verifyOnly As Boolean, ByVal usable As Variant)
    Dim view As SldWorks.View, annotation As SldWorks.Annotation, note As SldWorks.Note
    Dim geometry As Variant, bounds As Variant, position As Variant, anchor As Variant, target As Variant
    Dim leader As Variant, obstacle As Variant, occupied As Object
    Dim i As Long
    ReDim geometry(UBound(views))
    Set occupied = CreateObject("Scripting.Dictionary")
    For i = 0 To UBound(views)
        Set view = views(i)
        If Not view Is Nothing Then geometry(i) = view.GetOutline
    Next i
    For i = 0 To UBound(views)
        Set view = views(i)
        If Not view Is Nothing Then
            Set annotation = view.GetFirstAnnotation3
            Do While Not annotation Is Nothing
                If annotation.GetType = swAnnotationType_e.swNote And _
                   annotation.Visible <> swAnnotationVisibilityState_e.swAnnotationHidden Then
                    Set note = annotation.GetSpecificAnnotation
                    If note.IsBomBalloon And annotation.IsDangling = False Then
                        bounds = PDFBalloonBounds(annotation)
                        If Not verifyOnly Then
                            position = annotation.GetPosition
                            anchor = position
                            If annotation.GetLeaderCount > 0 Then
                                leader = annotation.GetLeaderPointsAtIndex(0)
                                If Not IsArray(leader) Then Err.Raise PDF_ERROR, "PDFClearAssemblyBalloons", "Cannot read balloon leader."
                                anchor = Array(leader(UBound(leader) - 2), leader(UBound(leader) - 1))
                            End If
                            target = PDFFindBalloonPosition(geometry(i), geometry, occupied, bounds, anchor, usable)
                            If annotation.SetPosition2(position(0) + target(0) - (bounds(0) + bounds(2)) / 2#, _
                                                      position(1) + target(1) - (bounds(1) + bounds(3)) / 2#, position(2)) = False Then
                                Err.Raise PDF_ERROR, "PDFClearAssemblyBalloons", "Cannot move BOM balloon outside the assembly."
                            End If
                            bounds = PDFBalloonBounds(annotation)
                        End If
                        If IsArray(usable) Then
                            If bounds(0) < usable(0) Or bounds(1) < usable(1) Or _
                               bounds(2) > usable(2) Or bounds(3) > usable(3) Then
                                Err.Raise PDF_ERROR, "PDFClearAssemblyBalloons", "BOM balloon " & note.GetBomBalloonText(True) & _
                                    " is outside the reserved drawing area in " & view.Name & "; verify=" & verifyOnly & _
                                    "; bounds=" & Join(Array(bounds(0), bounds(1), bounds(2), bounds(3)), ",")
                            End If
                        End If
                        For Each obstacle In geometry
                            If IsArray(obstacle) Then
                                If PDFBoxesOverlap(bounds, obstacle, 0.15 * INCH) Then
                                    Err.Raise PDF_ERROR, "PDFClearAssemblyBalloons", "BOM balloon " & note.GetBomBalloonText(True) & " overlaps assembly geometry in " & view.Name
                                End If
                            End If
                        Next obstacle
                        For Each obstacle In occupied.Items
                            If PDFBoxesOverlap(bounds, obstacle, 0.15 * INCH) Then
                                Err.Raise PDF_ERROR, "PDFClearAssemblyBalloons", "BOM balloon circles are too close in " & view.Name
                            End If
                        Next obstacle
                        occupied.Add occupied.Count, bounds
                    End If
                End If
                Set annotation = annotation.GetNext3
            Loop
        End If
    Next i
End Sub

Public Sub PDFClearPrintedBalloons(ByVal drawing As SldWorks.DrawingDoc, ByVal verifyOnly As Boolean)
    Dim sheets As Variant, sheetName As Variant, views As Variant, tables As Variant, usable As Variant
    Dim sheet As SldWorks.Sheet, view As SldWorks.View, sheetView As SldWorks.View
    Dim table As SldWorks.TableAnnotation
    Dim width As Double, height As Double, top As Double, count As Long, i As Long
    sheets = drawing.GetSheetNames
    For Each sheetName In sheets
        If drawing.ActivateSheet(CStr(sheetName)) = False Then Err.Raise PDF_ERROR, "PDFClearPrintedBalloons", "Cannot activate sheet for final balloon checks."
        Set sheet = drawing.GetCurrentSheet
        sheet.GetSize width, height
        top = height - 0.85 * INCH
        Set sheetView = drawing.GetFirstView
        tables = sheetView.GetTableAnnotations
        If IsArray(tables) Then
            For i = LBound(tables) To UBound(tables)
                Set table = tables(i)
                If table.Type = swTableAnnotationType_e.swTableAnnotation_BillOfMaterials Then top = top - 1.6 * INCH
            Next i
        End If
        usable = Array(0.85 * INCH, 0.85 * INCH, width - 2.5 * INCH, top)
        Set view = sheetView.GetNextView
        count = 0
        ReDim views(0)
        Do While Not view Is Nothing
            ReDim Preserve views(count)
            Set views(count) = view
            count = count + 1
            Set view = view.GetNextView
        Loop
        If count > 0 Then PDFClearAssemblyBalloons views, verifyOnly, usable
    Next sheetName
End Sub

Public Function PDFHidePrintArtifacts(ByVal drawing As SldWorks.DrawingDoc, ByVal model As SldWorks.ModelDoc2) As Long
    Dim view As SldWorks.View
    Dim annotation As SldWorks.Annotation
    If model.GetUserPreferenceToggle(swUserPreferenceToggle_e.swDisplaySketches) Then
        Call model.SetUserPreferenceToggle(swUserPreferenceToggle_e.swDisplaySketches, False)
    End If
    If model.GetUserPreferenceToggle(swUserPreferenceToggle_e.swDisplaySketches) Then
        Err.Raise PDF_ERROR, "PDFHidePrintArtifacts", "Sketch display is still enabled."
    End If
    Set view = drawing.GetFirstView
    Do While Not view Is Nothing
        Set annotation = view.GetFirstAnnotation3
        Do While Not annotation Is Nothing
            If annotation.GetType = swAnnotationType_e.swDisplayDimension Or _
               annotation.GetType = swAnnotationType_e.swCenterMarkSym Then
                If annotation.IsDangling And annotation.Visible <> swAnnotationVisibilityState_e.swAnnotationHidden Then
                    annotation.Visible = swAnnotationVisibilityState_e.swAnnotationHidden
                    If annotation.Visible <> swAnnotationVisibilityState_e.swAnnotationHidden Then
                        Err.Raise PDF_ERROR, "PDFHidePrintArtifacts", "Cannot hide a dangling dimension or center mark in view " & view.Name
                    End If
                    PDFHidePrintArtifacts = PDFHidePrintArtifacts + 1
                End If
            End If
            Set annotation = annotation.GetNext3
        Loop
        Set view = view.GetNextView
    Loop
End Function

Public Sub PDFPositionBaseplateBalloon(ByVal view As SldWorks.View, ByVal annotation As SldWorks.Annotation, _
                                      ByVal component As SldWorks.Component2)
    Dim math As SldWorks.MathUtility, point As SldWorks.MathPoint
    Dim edge As SldWorks.Edge, curve As SldWorks.Curve
    Dim startVertex As SldWorks.Vertex, endVertex As SldWorks.Vertex
    Dim edges As Variant, start As Variant, finish As Variant, position As Variant, target As Variant
    Dim midpoint(2) As Double, i As Long
    Set math = Application.SldWorks.GetMathUtility
    edges = view.GetVisibleEntities2(component, swViewEntityType_e.swViewEntityType_Edge)
    If Not IsArray(edges) Then Err.Raise PDF_ERROR, "PDFPositionBaseplateBalloon", "No visible baseplate edges."
    For i = LBound(edges) To UBound(edges)
        Set edge = edges(i)
        Set curve = edge.GetCurve
        If curve.IsLine Then
            Set startVertex = edge.GetStartVertex
            Set endVertex = edge.GetEndVertex
            If Not startVertex Is Nothing And Not endVertex Is Nothing Then
                start = startVertex.GetPoint
                finish = endVertex.GetPoint
                midpoint(0) = (start(0) + finish(0)) / 2#
                midpoint(1) = (start(1) + finish(1)) / 2#
                midpoint(2) = (start(2) + finish(2)) / 2#
                Set point = math.CreatePoint(midpoint)
                Set point = point.MultiplyTransform(component.GetTotalTransform(True))
                Set point = point.MultiplyTransform(view.ModelToViewTransform)
                position = point.ArrayData
                If IsEmpty(target) Then
                    target = position
                ElseIf position(0) < target(0) Then
                    target = position
                End If
            End If
        End If
    Next i
    If IsEmpty(target) Then Err.Raise PDF_ERROR, "PDFPositionBaseplateBalloon", "No straight baseplate edge found for a clear leader."
    ' Use the plate's leftmost visible edge, away from gussets near the other end.
    If annotation.SetLeaderAttachmentPointAtIndex(0, target(0), target(1), 0#) = False Then
        Err.Raise PDF_ERROR, "PDFPositionBaseplateBalloon", "Cannot position the baseplate leader."
    End If
    If annotation.SetPosition2(target(0) - INCH, target(1) - INCH, 0#) = False Then
        Err.Raise PDF_ERROR, "PDFPositionBaseplateBalloon", "Cannot position the baseplate balloon."
    End If
End Sub

Public Sub PDFPositionExplodedBaseplates(ByVal view As SldWorks.View)
    Dim annotation As SldWorks.Annotation, note As SldWorks.Note
    Dim entities As Variant, entity As SldWorks.Entity, component As SldWorks.Component2
    Set annotation = view.GetFirstAnnotation3
    Do While Not annotation Is Nothing
        If annotation.GetType = swAnnotationType_e.swNote And _
           annotation.Visible <> swAnnotationVisibilityState_e.swAnnotationHidden Then
            Set note = annotation.GetSpecificAnnotation
            If note.IsBomBalloon And annotation.IsDangling = False Then
                entities = annotation.GetAttachedEntities3
                If Not IsArray(entities) Then Err.Raise PDF_ERROR, "PDFPositionExplodedBaseplates", "BOM balloon has no attached entity."
                Set entity = entities(LBound(entities))
                Set component = entity.GetComponent
                If component Is Nothing Then Err.Raise PDF_ERROR, "PDFPositionExplodedBaseplates", "BOM balloon has no attached component."
                If InStr(1, component.GetPathName, "-BP-", vbTextCompare) > 0 Then
                    PDFPositionBaseplateBalloon view, annotation, component
                End If
            End If
        End If
        Set annotation = annotation.GetNext3
    Loop
End Sub

Public Function PDFRefreshExplodedBalloons(ByVal drawing As SldWorks.DrawingDoc, _
                                          ByVal model As SldWorks.ModelDoc2, ByVal view As SldWorks.View) As Long
    Dim options As SldWorks.AutoBalloonOptions
    Dim result As Variant, annotation As SldWorks.Annotation, note As SldWorks.Note
    Dim tables As Variant, table As SldWorks.TableAnnotation, bom As SldWorks.BomTableAnnotation
    Dim sheetView As SldWorks.View
    Dim components As Variant, component As SldWorks.Component2
    Dim entities As Variant, entity As SldWorks.Entity
    Dim expected As Object, found As Object
    Dim i As Long, row As Long, count As Long
    Dim itemNumber As String, partNumber As String, key As Variant
    Set expected = CreateObject("Scripting.Dictionary")
    Set found = CreateObject("Scripting.Dictionary")
    Set sheetView = drawing.GetFirstView
    tables = sheetView.GetTableAnnotations
    If IsArray(tables) Then
        For i = LBound(tables) To UBound(tables)
            Set table = tables(i)
            If table.Type = swTableAnnotationType_e.swTableAnnotation_BillOfMaterials Then
                Set bom = table
                For row = 0 To table.RowCount - 1
                    itemNumber = ""
                    partNumber = ""
                    count = bom.GetComponentsCount2(row, view.ReferencedConfiguration, itemNumber, partNumber)
                    If count > 0 And itemNumber <> "" Then
                        components = bom.GetComponents2(row, view.ReferencedConfiguration)
                        If Not IsArray(components) Then Err.Raise PDF_ERROR, "PDFRefreshExplodedBalloons", "Cannot read BOM components."
                        Set component = components(LBound(components))
                        expected(itemNumber & "|" & LCase$(component.GetPathName)) = True
                    End If
                Next row
            End If
        Next i
    End If
    If expected.Count = 0 Then Err.Raise PDF_ERROR, "PDFRefreshExplodedBalloons", "No component rows found in the exploded sheet BOM."
    model.ClearSelection2 True
    If model.Extension.SelectByID2(view.Name, "DRAWINGVIEW", 0#, 0#, 0#, False, 0, Nothing, 0) = False Then
        Err.Raise PDF_ERROR, "PDFRefreshExplodedBalloons", "Cannot select the exploded view."
    End If
    Set options = drawing.CreateAutoBalloonOptions
    options.Layout = swBalloonLayoutType_e.swDetailingBalloonLayout_Square
    options.IgnoreMultiple = True
    options.LeaderAttachmentToFaces = True
    options.Style = swBalloonStyle_e.swBS_Circular
    options.Size = swBalloonFit_e.swBF_2Chars
    options.UpperTextContent = swBalloonTextContent_e.swBalloonTextItemNumber
    options.ItemOrder = swBalloonItemNumbersOrder_e.swBalloonItemNumbers_DoNotChangeItemNumbers
    options.InsertMagneticLine = False
    options.EditBalloons = True
    options.EditBalloonOption = swEditBalloonOption_e.swEditBalloonOption_Replace
    result = drawing.AutoBalloon5(options)
    model.ClearSelection2 True
    If Not IsArray(result) Then Err.Raise PDF_ERROR, "PDFRefreshExplodedBalloons", "SolidWorks could not regenerate BOM balloons."
    Set annotation = view.GetFirstAnnotation3
    Do While Not annotation Is Nothing
        If annotation.GetType = swAnnotationType_e.swNote Then
            Set note = annotation.GetSpecificAnnotation
            If note.IsBomBalloon Then
                If annotation.IsDangling Then
                    annotation.Visible = swAnnotationVisibilityState_e.swAnnotationHidden
                    If annotation.Visible <> swAnnotationVisibilityState_e.swAnnotationHidden Then
                        Err.Raise PDF_ERROR, "PDFRefreshExplodedBalloons", "Cannot hide a stale BOM balloon."
                    End If
                Else
                    annotation.Visible = swAnnotationVisibilityState_e.swAnnotationVisible
                    If annotation.Visible <> swAnnotationVisibilityState_e.swAnnotationVisible Then
                        Err.Raise PDF_ERROR, "PDFRefreshExplodedBalloons", "Cannot show a valid BOM balloon."
                    End If
                    entities = annotation.GetAttachedEntities3
                    If Not IsArray(entities) Then Err.Raise PDF_ERROR, "PDFRefreshExplodedBalloons", "BOM balloon has no attached entity."
                    Set entity = entities(LBound(entities))
                    Set component = entity.GetComponent
                    If component Is Nothing Then Err.Raise PDF_ERROR, "PDFRefreshExplodedBalloons", "BOM balloon has no attached component."
                    key = note.GetBomBalloonText(True) & "|" & LCase$(component.GetPathName)
                    If Not expected.Exists(key) Then Err.Raise PDF_ERROR, "PDFRefreshExplodedBalloons", "Balloon/component does not match the BOM: " & note.GetBomBalloonText(True)
                    found(key) = True
                End If
            End If
        End If
        Set annotation = annotation.GetNext3
    Loop
    For Each key In expected.Keys
        If Not found.Exists(key) Then Err.Raise PDF_ERROR, "PDFRefreshExplodedBalloons", "Missing BOM-linked balloon: " & key
    Next key
    PDFRefreshExplodedBalloons = found.Count
End Function

Public Sub PDFLayoutViews(ByVal drawing As SldWorks.DrawingDoc, ByVal model As SldWorks.ModelDoc2, _
                          ByVal view1 As SldWorks.View, ByVal view2 As SldWorks.View, _
                          ByVal view3 As SldWorks.View, ByVal kind As String)
    Dim views As Variant, widths As Variant, heights As Variant
    Dim scales As Variant, centersX As Variant, centersY As Variant
    Dim measuredBounds As Variant
    Dim sheet As SldWorks.Sheet, view As SldWorks.View
    Dim bounds As Variant, anchor As Variant, newPosition(1) As Double
    Dim sheetWidth As Double, sheetHeight As Double, currentScale As Double
    Dim i As Long
    Dim scaleChanged As Boolean
    Set sheet = drawing.GetCurrentSheet
    sheet.GetSize sheetWidth, sheetHeight
    views = Array(view1, view2, view3)
    widths = Array(0#, 0#, 0#)
    heights = Array(0#, 0#, 0#)
    measuredBounds = Array(Empty, Empty, Empty)
    For i = 0 To 2
        Set view = views(i)
        If Not view Is Nothing Then
            view.RemoveAlignment
            view.PositionLocked = False
            currentScale = view.ScaleDecimal
            If currentScale <= 0 Then Err.Raise PDF_ERROR, "PDFLayoutViews", "Invalid view scale: " & view.Name
            bounds = PDFViewBounds(view)
            measuredBounds(i) = bounds
            widths(i) = (bounds(2) - bounds(0)) / currentScale
            heights(i) = (bounds(3) - bounds(1)) / currentScale
        End If
    Next i
    PDFPlanViews kind, widths, heights, sheetWidth, sheetHeight, scales, centersX, centersY
    For i = 0 To 2
        Set view = views(i)
        If Not view Is Nothing Then
            view.UseSheetScale = False
            view.UseParentScale = False
            If Abs(view.ScaleDecimal - scales(i)) > 0.000000001 Then
                view.ScaleDecimal = scales(i)
                scaleChanged = True
            End If
        End If
    Next i
    For i = 0 To 2
        If scales(i) > 0 Then
            If sheet.SetScale(1#, 1# / scales(i), False, False) = False Then
                Err.Raise PDF_ERROR, "PDFLayoutViews", "Cannot update the sheet scale."
            End If
            Exit For
        End If
    Next i
    If scaleChanged Then
        If model.EditRebuild3 = False Then Err.Raise PDF_ERROR, "PDFLayoutViews", "Drawing rebuild failed after scaling."
    End If
    For i = 0 To 2
        Set view = views(i)
        If Not view Is Nothing Then
            If scaleChanged Then
                bounds = PDFViewBounds(view)
            Else
                bounds = measuredBounds(i)
            End If
            anchor = view.Position
            newPosition(0) = anchor(0) + centersX(i) - (bounds(0) + bounds(2)) / 2#
            newPosition(1) = anchor(1) + centersY(i) - (bounds(1) + bounds(3)) / 2#
            view.Position = newPosition
        End If
    Next i
    If kind = "exploded" Then
        If InStr(1, view1.GetReferencedModelName, "\1-S", vbTextCompare) > 0 Then
            PDFPositionExplodedBaseplates view1
        End If
    End If
    If kind <> "component" Then PDFClearAssemblyBalloons views, False, Empty
End Sub

Public Function PDFPrepareSheet(ByVal drawing As SldWorks.DrawingDoc, ByVal model As SldWorks.ModelDoc2, _
                                 ByVal view1 As SldWorks.View, ByVal view2 As SldWorks.View, _
                                 ByVal view3 As SldWorks.View, ByVal kind As String) As String
    Dim views As Variant, view As SldWorks.View
    Dim annotation As SldWorks.Annotation, dimension As SldWorks.DisplayDimension
    Dim selection As SldWorks.SelectData
    Dim i As Long, selected As Long, jogWarnings As Long, hidden As Long, balloons As Long, attempt As Long
    hidden = PDFHidePrintArtifacts(drawing, model)
    If model.ForceRebuild3(False) = False Then Err.Raise PDF_ERROR, "PDFPrepareSheet", "Drawing rebuild failed."
    hidden = hidden + PDFHidePrintArtifacts(drawing, model)
    If kind = "exploded" Then
        If InStr(1, view1.GetReferencedModelName, "\1-S", vbTextCompare) > 0 Then
            balloons = PDFRefreshExplodedBalloons(drawing, model, view1)
        End If
    End If
    views = Array(view1, view2, view3)
    If kind <> "component" Then PDFClearAssemblyBalloons views, False, Empty
    PDFLayoutViews drawing, model, view1, view2, view3, kind
    model.ClearSelection2 True
    Set selection = model.SelectionManager.CreateSelectData
    views = Array(view1, view2, view3)
    For i = 0 To 2
        Set view = views(i)
        If Not view Is Nothing Then
            Set annotation = view.GetFirstAnnotation3
            Do While Not annotation Is Nothing
                If annotation.GetType = swAnnotationType_e.swDisplayDimension And _
                   annotation.Visible <> swAnnotationVisibilityState_e.swAnnotationHidden Then
                    Set dimension = annotation.GetSpecificAnnotation
                    If Not dimension Is Nothing Then
                        If dimension.Type2 = swDimensionType_e.swOrdinateDimension Then
                            If dimension.AutoJogOrdinate() = False Then jogWarnings = jogWarnings + 1
                        End If
                        If annotation.Select3(True, selection) = False Then
                            Err.Raise PDF_ERROR, "PDFPrepareSheet", "Cannot select a dimension in view " & view.Name
                        End If
                        selected = selected + 1
                    End If
                End If
                Set annotation = annotation.GetNext3
            Loop
        End If
    Next i
    If selected > 0 Then
        If model.Extension.AlignDimensions(swAlignDimensionType_e.swAlignDimensionType_AutoArrange, 0.003) = False Then
            Err.Raise PDF_ERROR, "PDFPrepareSheet", "Cannot arrange drawing dimensions."
        End If
    End If
    model.ClearSelection2 True
    ' Dimensions can extend beyond geometry; refit their final positions before export.
    For attempt = 1 To 4
        PDFLayoutViews drawing, model, view1, view2, view3, kind
        If model.EditRebuild3 = False Then Err.Raise PDF_ERROR, "PDFPrepareSheet", "Final drawing rebuild failed."
        If kind = "component" Then Exit For
        If PDFViewsFit(drawing, views) Then Exit For
    Next attempt
    If attempt > 4 Then Err.Raise PDF_ERROR, "PDFPrepareSheet", "Cannot fit views and balloons inside the reserved drawing area."
    PDFVerifyViews drawing, view1, view2, view3
    PDFPrepareSheet = "Layout=" & kind & "; dimensions=" & selected & "; auto-jog warnings=" & jogWarnings & _
        "; hidden dangling dimensions/center marks=" & hidden & "; sketches hidden; BOM balloons=" & balloons
End Function

Public Function PDFViewsFit(ByVal drawing As SldWorks.DrawingDoc, ByVal views As Variant) As Boolean
    Dim sheet As SldWorks.Sheet, view As SldWorks.View
    Dim width As Double, height As Double, bounds As Variant, i As Long
    Set sheet = drawing.GetCurrentSheet
    sheet.GetSize width, height
    PDFViewsFit = True
    For i = 0 To UBound(views)
        Set view = views(i)
        If Not view Is Nothing Then
            bounds = PDFViewBounds(view)
            If bounds(0) < 0.85 * INCH Or bounds(1) < 0.85 * INCH Or _
               bounds(2) > width - 2.5 * INCH Or bounds(3) > height - 0.85 * INCH Then
                PDFViewsFit = False
            End If
        End If
    Next i
End Function

Public Sub PDFVerifyViews(ByVal drawing As SldWorks.DrawingDoc, ByVal view1 As SldWorks.View, _
                          ByVal view2 As SldWorks.View, ByVal view3 As SldWorks.View)
    Dim sheet As SldWorks.Sheet, view As SldWorks.View
    Dim annotation As SldWorks.Annotation
    Dim views As Variant, bounds As Variant, otherBounds As Variant, measuredBounds As Variant
    Dim other As SldWorks.View
    Dim width As Double, height As Double, i As Long, j As Long
    Set sheet = drawing.GetCurrentSheet
    sheet.GetSize width, height
    views = Array(view1, view2, view3)
    PDFClearAssemblyBalloons views, True, Empty
    measuredBounds = Array(Empty, Empty, Empty)
    For i = 0 To 2
        Set view = views(i)
        If Not view Is Nothing Then measuredBounds(i) = PDFViewBounds(view)
    Next i
    For i = 0 To 2
        Set view = views(i)
        If Not view Is Nothing Then
            Set annotation = view.GetFirstAnnotation3
            Do While Not annotation Is Nothing
                If annotation.GetType = swAnnotationType_e.swDisplayDimension Or _
                   annotation.GetType = swAnnotationType_e.swCenterMarkSym Then
                    If annotation.IsDangling And annotation.Visible <> swAnnotationVisibilityState_e.swAnnotationHidden Then
                        Err.Raise PDF_ERROR, "PDFVerifyViews", "A dangling dimension or center mark is still visible in view " & view.Name
                    End If
                End If
                Set annotation = annotation.GetNext3
            Loop
            bounds = measuredBounds(i)
            If bounds(0) < 0.5 * INCH Or bounds(1) < 0.5 * INCH Or _
               bounds(2) > width - 2.15 * INCH Or bounds(3) > height - 0.5 * INCH Then
                Err.Raise PDF_ERROR, "PDFVerifyViews", "View or annotation is outside the usable sheet area: " & view.Name
            End If
            For j = i + 1 To 2
                Set other = views(j)
                If Not other Is Nothing Then
                    otherBounds = measuredBounds(j)
                    If PDFMin(bounds(2), otherBounds(2)) > PDFMax(bounds(0), otherBounds(0)) + 0.0001 And _
                       PDFMin(bounds(3), otherBounds(3)) > PDFMax(bounds(1), otherBounds(1)) + 0.0001 Then
                        Err.Raise PDF_ERROR, "PDFVerifyViews", "View/annotation bounds overlap: " & view.Name & " and " & other.Name
                    End If
                End If
            Next j
        End If
    Next i
End Sub

Public Sub PDFExportDrawing(ByVal app As SldWorks.SldWorks, ByVal model As SldWorks.ModelDoc2, ByVal path As String)
    Dim data As SldWorks.ExportPdfData
    Dim drawing As SldWorks.DrawingDoc
    Dim sheets As Variant
    Dim errors As Long, warnings As Long
    If model.GetUserPreferenceToggle(swUserPreferenceToggle_e.swDisplaySketches) Then
        Err.Raise PDF_ERROR, "PDFExportDrawing", "Sketch display must be disabled before exporting."
    End If
    Set drawing = model
    ' Reactivating sheets refreshes printed balloon geometry after view fitting.
    PDFClearPrintedBalloons drawing, False
    If model.Save3(swSaveAsOptions_e.swSaveAsOptions_Silent, errors, warnings) = False Or errors <> 0 Then
        Err.Raise PDF_ERROR, "PDFExportDrawing", "Cannot save drawing; errors=" & errors & ", warnings=" & warnings
    End If
    If warnings <> 0 Then Err.Raise PDF_ERROR, "PDFExportDrawing", "Drawing save warnings require review: " & warnings
    PDFClearPrintedBalloons drawing, True
    Set data = app.GetExportFileData(swExportDataFileType_e.swExportPdfData)
    sheets = drawing.GetSheetNames
    data.ViewPdfAfterSaving = False
    If data.SetSheets(swExportDataSheetsToExport_e.swExportData_ExportSpecifiedSheets, sheets) = False Then
        Err.Raise PDF_ERROR, "PDFExportDrawing", "Cannot select all drawing sheets for PDF export."
    End If
    errors = 0
    warnings = 0
    If model.Extension.SaveAs(path, swSaveAsVersion_e.swSaveAsCurrentVersion, _
                             swSaveAsOptions_e.swSaveAsOptions_Silent, data, errors, warnings) = False Or errors <> 0 Then
        Err.Raise PDF_ERROR, "PDFExportDrawing", "PDF export failed; errors=" & errors & ", warnings=" & warnings & vbCrLf & path
    End If
    If warnings <> 0 Then Err.Raise PDF_ERROR, "PDFExportDrawing", "PDF export warnings require review: " & warnings & vbCrLf & path
    If Dir$(path) = "" Then Err.Raise PDF_ERROR, "PDFExportDrawing", "PDF was not created: " & path
    If FileLen(path) = 0 Then Err.Raise PDF_ERROR, "PDFExportDrawing", "PDF is empty: " & path
End Sub
