Attribute VB_Name = "ui_quote_calculator"
Option Explicit

Public Function UICalculateQuote(ByVal walls As Long, ByVal orientation As String, _
    ByVal returnLabel As String, ByVal material As String, ByVal thicknessLabel As String, _
    ByVal length As Double, ByVal width As Double, ByVal height As Double, _
    Optional ByVal pricePerPound As Double = 0.7) As UIQuoteResult
    Dim result As UIQuoteResult
    Dim available As Double
    UIValidateSelections walls, orientation, returnLabel, material, thicknessLabel
    UIValidateDimensions length, width, height
    If pricePerPound <= 0 Then
        Err.Raise vbObjectError + 2400, "UICalculateQuote", "Material price must be greater than zero dollars per lb."
    End If
    Set result = New UIQuoteResult
    result.Density = 7850 * (0.0254 ^ 3) / 0.45359237
    result.PricePerPound = pricePerPound
    result.WallCount = walls
    result.Orientation = orientation
    result.ReturnLabel = returnLabel
    result.Material = material
    result.Thickness = IIf(thicknessLabel = "3/16 inch", 0.1875, 0.25)
    result.OverallLength = length
    result.OverallWidth = width
    result.Height = height
    Select Case walls
        Case 1: available = length
        Case 2: available = length - height / 2
        Case 3, 4: available = length - height
    End Select
    ' Workbook N10/N11 sizing and assembly pattern counts deliberately use different rules.
    result.CenterSizingSegments = -Int(-available / 120)
    result.SideSizingSegments = -Int(-width / 120)
    result.CenterLength = available / result.CenterSizingSegments
    result.SideLength = width / result.SideSizingSegments
    If result.SideSizingSegments > 1 Then
        result.SideLength = result.SideLength + result.Thickness + _
            result.Thickness * (result.SideSizingSegments - 2) / result.SideSizingSegments
    End If
    result.CenterPatternCount = Int((available - 1) / 120) + 1
    result.SidePlantersPerWall = Int((width - 1) / 120) + 1
    result.CenterWallRuns = IIf(walls = 4, 2, 1)
    Select Case walls
        Case 1: result.SideWallRuns = 0
        Case 2: result.SideWallRuns = 1
        Case 3, 4: result.SideWallRuns = 2
    End Select
    result.CenterGussets = UIQuoteGussets(result.CenterLength)
    result.SideGussets = UIQuoteGussets(result.SideLength)
    result.SideMate = result.CenterLength / 2 + height / 2
    result.SideMateRight = result.SideMate + (result.CenterPatternCount - 1) * result.CenterLength
    If result.CenterPatternCount <> result.CenterSizingSegments Then
        result.Warnings = "Center sizing/pattern counts disagree: quantities require review." & vbCrLf
    End If
    If result.SideWallRuns > 0 And result.SidePlantersPerWall <> result.SideSizingSegments Then
        result.Warnings = result.Warnings & "Side sizing/pattern counts disagree: quantities require review." & vbCrLf
    End If
    UIQuoteComponents result
    Set UICalculateQuote = result
End Function

Public Function UIQuoteParsePrice(ByVal text As String) As Double
    On Error GoTo InvalidPrice
    UIQuoteParsePrice = UIParseInches(text)
    Exit Function
InvalidPrice:
    Err.Raise vbObjectError + 2400, "UIQuoteParsePrice", _
        "Enter a positive decimal material price in $/lb, for example 0.70 (no currency symbol or commas)."
End Function

Private Function UIQuoteFlatHeight(ByVal height As Double, ByVal thickness As Double, _
    ByVal returnLabel As String) As Double
    Dim singleOffset As Double
    Dim doubleOffset As Double
    ' Preserve measured offsets independently: the supplied 3/16 double value is rounded differently.
    If thickness = 0.25 Then
        singleOffset = 0.5956
        doubleOffset = 1.1912
    ElseIf thickness = 0.1875 Then
        singleOffset = 0.6798
        doubleOffset = 1.3595
    Else
        Err.Raise vbObjectError + 2400, "UIQuoteFlatHeight", "Unsupported material thickness."
    End If
    Select Case returnLabel
        Case "No return": UIQuoteFlatHeight = height
        Case "Single return 1 inch": UIQuoteFlatHeight = height + singleOffset
        Case "Single return 2 inches": UIQuoteFlatHeight = height + 1 + singleOffset
        Case "Double return": UIQuoteFlatHeight = height + doubleOffset
        Case Else: Err.Raise vbObjectError + 2400, "UIQuoteFlatHeight", "Unsupported return type."
    End Select
End Function

Private Function UIQuotePart(ByVal result As UIQuoteResult, ByVal name As String, _
    ByVal planterType As String, ByVal length As Double, ByVal width As Double, _
    ByVal topWidth As Double, ByVal quantity As Double, ByVal dimensionRule As String) As UIQuoteComponent
    Dim part As UIQuoteComponent
    Set part = New UIQuoteComponent
    If length <= 0 Or width <= 0 Or quantity < 0 Then
        Err.Raise vbObjectError + 2400, "UIQuotePart", "Invalid component dimensions or quantity: " & name
    End If
    part.Name = name
    part.PlanterType = planterType
    part.Material = result.Material
    part.Thickness = result.Thickness
    part.Length = length
    part.Width = width
    part.TopWidth = topWidth
    part.DimensionRule = dimensionRule
    If topWidth > 0 Then
        part.Shape = "Trapezoid"
        part.AreaRule = "(top width + bottom width) / 2 * height"
        part.Area = (topWidth + width) / 2 * length
    Else
        part.Shape = "Rectangle"
        part.AreaRule = "length * width"
        part.Area = length * width
    End If
    part.Quantity = quantity
    part.TotalArea = part.Area * quantity
    part.Weight = part.Area * part.Thickness * result.Density
    part.TotalWeight = part.Weight * quantity
    part.MaterialCost = part.Weight * result.PricePerPound
    part.TotalMaterialCost = part.TotalWeight * result.PricePerPound
    Set UIQuotePart = part
End Function

Private Sub UIQuoteComponents(ByVal result As UIQuoteResult)
    Dim centers As Double
    Dim sides As Double
    Dim parts(0 To 7) As UIQuoteComponent
    Dim i As Long
    centers = result.CenterPatternCount * result.CenterWallRuns
    sides = result.SidePlantersPerWall * result.SideWallRuns
    result.FlatHeight = UIQuoteFlatHeight(result.Height, result.Thickness, result.ReturnLabel)
    result.SideBlankLength = result.SideLength - 2 * result.Thickness
    result.GussetHeight = result.Height - 2.1 - result.Thickness
    result.GussetBottomWidth = result.Height / 2 - result.Thickness - 0.5
    result.GussetTopWidth = result.GussetBottomWidth / 3
    Set parts(0) = UIQuotePart(result, "Center Planter Wall", "Center", result.CenterLength, result.FlatHeight, 0, centers, "CENTER_PLANTER_LENGTH x calibrated flat height")
    Set parts(1) = UIQuotePart(result, "Center Planter Base Plate", "Center", result.CenterLength, result.Height / 2, 0, centers, "CENTER_PLANTER_LENGTH x H/2")
    Set parts(2) = UIQuotePart(result, "Side Planter Wall", "Side", result.SideBlankLength, result.FlatHeight, 0, sides, "(SIDE_PLANTER_LENGTH - 2*T) x calibrated flat height")
    Set parts(3) = UIQuotePart(result, "Side Planter Base Plate", "Side", result.SideBlankLength, result.Height / 2, 0, sides, "(SIDE_PLANTER_LENGTH - 2*T) x H/2")
    Set parts(4) = UIQuotePart(result, "leftsidewall", "Side ends", result.Height / 2, result.FlatHeight, 0, result.SideWallRuns, "H/2 x calibrated flat height; one per side run")
    Set parts(5) = UIQuotePart(result, "rightsidewall", "Side ends", result.Height / 2, result.FlatHeight, 0, result.SideWallRuns, "H/2 x calibrated flat height; one per side run")
    Set parts(6) = UIQuotePart(result, "Mounting Gusset", "Shared", result.GussetHeight, result.GussetBottomWidth, result.GussetTopWidth, 2 * (centers + sides), "Height=H-2.1-T; bottom=H/2-T-0.5; top=bottom/3; qty=2 per planter")
    Set parts(7) = UIQuotePart(result, "Center Gusset", "Shared", result.GussetHeight, result.GussetBottomWidth, result.GussetTopWidth, centers * (result.CenterGussets - 1) + sides * (result.SideGussets - 1), "Same profile as mounting; qty=centers*(CENTER_GUSSETS-1)+sides*(SIDE_GUSSETS-1)")
    result.Components = parts
    result.HasValidTotals = (Len(result.Warnings) = 0)
    If result.HasValidTotals Then
        For i = 0 To 7
            result.TotalArea = result.TotalArea + parts(i).TotalArea
            result.TotalWeight = result.TotalWeight + parts(i).TotalWeight
            result.TotalMaterialCost = result.TotalMaterialCost + parts(i).TotalMaterialCost
        Next i
    End If
End Sub

Private Function UIQuoteGussets(ByVal length As Double) As Long
    Select Case length
        Case Is <= 24: UIQuoteGussets = 1
        Case Is <= 48: UIQuoteGussets = 2
        Case Is <= 72: UIQuoteGussets = 3
        Case Is <= 96: UIQuoteGussets = 4
        Case Else: UIQuoteGussets = 5
    End Select
End Function

Public Function UIQuoteReport(ByVal result As UIQuoteResult) As String
    Dim text As String
    Dim centers As Double
    Dim sides As Double
    Dim part As UIQuoteComponent
    Dim i As Long
    centers = result.CenterPatternCount * result.CenterWallRuns
    sides = result.SidePlantersPerWall * result.SideWallRuns
    text = "INPUT-ONLY COMPONENT QUOTE - formula version 3" & vbCrLf & _
        "No model opened, edited, rebuilt or measured." & vbCrLf & vbCrLf & _
        "Overall walls: " & result.WallCount & "; orientation: " & result.Orientation & vbCrLf & _
        "Overall L x W x H (in): " & UIInchesText(result.OverallLength) & " x " & _
        UIInchesText(result.OverallWidth) & " x " & UIInchesText(result.Height) & vbCrLf & _
        "Material: " & result.Material & "; thickness (in): " & UIInchesText(result.Thickness) & vbCrLf & _
        "Return: " & result.ReturnLabel & vbCrLf & _
        "Estimating density: " & UIInchesText(result.Density) & " lb/in^3 (7850 kg/m^3 for both steels)." & vbCrLf & _
        "BORCON-specific density is unverified; shared steel density is an approved estimating assumption." & vbCrLf & _
        "Material price ($/lb): " & UIInchesText(result.PricePerPound) & vbCrLf & vbCrLf & _
        "Center wall runs: " & result.CenterWallRuns & vbCrLf & _
        "Center sizing segments per run (workbook): " & UIInchesText(result.CenterSizingSegments) & vbCrLf & _
        "CENTER_WALL_PATTERN_COUNT (assembly): " & UIInchesText(result.CenterPatternCount) & vbCrLf & _
        "CENTER_PLANTER_LENGTH (in): " & UIInchesText(result.CenterLength) & vbCrLf & _
        "CENTER_GUSSETS equation value: " & result.CenterGussets & vbCrLf & vbCrLf & _
        "Side wall runs: " & result.SideWallRuns & vbCrLf & _
        "Side sizing segments per run (workbook): " & UIInchesText(result.SideSizingSegments) & vbCrLf & _
        "SIDE_WALL_PLANTER_COUNT per side run (assembly): " & UIInchesText(result.SidePlantersPerWall) & vbCrLf & _
        "SIDE_PLANTER_LENGTH (in): " & UIInchesText(result.SideLength) & vbCrLf & _
        "SIDE_GUSSETS equation value: " & result.SideGussets & vbCrLf & vbCrLf & _
        "SIDE_WALL_MATE (in): " & UIInchesText(result.SideMate) & vbCrLf & _
        "SIDE_WALL_MATE_RIGHT (in): " & UIInchesText(result.SideMateRight) & vbCrLf & vbCrLf
    If Len(result.Warnings) > 0 Then
        text = text & "REVIEW REQUIRED:" & vbCrLf & result.Warnings & _
            "Total quantities, areas, weights and costs withheld until counts agree." & vbCrLf
    Else
        text = text & "Derived center planter total: " & UIInchesText(centers) & vbCrLf & _
            "Derived side planter total: " & UIInchesText(sides) & vbCrLf & _
            "Mounting gussets total: " & UIInchesText(2 * (centers + sides)) & vbCrLf & _
            "Center gussets total: " & UIInchesText(centers * (result.CenterGussets - 1) + sides * (result.SideGussets - 1)) & vbCrLf
    End If
    text = text & vbCrLf & "Flat wall height (in): " & UIInchesText(result.FlatHeight) & vbCrLf & _
        "Calibrated from rounded 48-inch measurements; other heights are predictions." & vbCrLf & _
        "COMPONENT BREAKDOWN - dimensions in inches; areas in square inches" & vbCrLf
    For i = LBound(result.Components) To UBound(result.Components)
        Set part = result.Components(i)
        text = text & vbCrLf & part.Name & " [" & part.PlanterType & "]" & vbCrLf & _
            "Material: " & part.Material & "; T=" & UIInchesText(part.Thickness) & vbCrLf & _
            "Rule: " & part.DimensionRule & vbCrLf & _
            "Length=" & UIInchesText(part.Length) & "; width/height=" & UIInchesText(part.Width)
        If part.TopWidth > 0 Then text = text & "; top=" & UIInchesText(part.TopWidth)
        text = text & vbCrLf & part.Shape & " area: " & part.AreaRule & vbCrLf & _
            "Area per part: " & UIInchesText(part.Area) & vbCrLf & _
            "Profile weight per part (lb): " & UIInchesText(part.Weight) & _
            "; material cost per part ($): " & FormatNumber(part.MaterialCost, 2) & vbCrLf
        If result.HasValidTotals Then
            text = text & "Qty: " & UIInchesText(part.Quantity) & "; total area: " & UIInchesText(part.TotalArea) & vbCrLf
            text = text & "Total profile weight (lb): " & UIInchesText(part.TotalWeight) & _
                "; total material cost ($): " & FormatNumber(part.TotalMaterialCost, 2) & vbCrLf
        Else
            text = text & "Qty, total area, weight and cost: WITHHELD - count mismatch" & vbCrLf
        End If
    Next i
    If result.HasValidTotals Then
        text = text & vbCrLf & "TOTAL PROFILE AREA (" & result.Material & ", T=" & _
            UIInchesText(result.Thickness) & "): " & UIInchesText(result.TotalArea) & _
            " in^2 / " & UIInchesText(result.TotalArea / 144) & " ft^2" & vbCrLf & _
            "TOTAL PROFILE WEIGHT (lb): " & UIInchesText(result.TotalWeight) & vbCrLf & _
            "TOTAL PROFILE MATERIAL COST ($): " & FormatNumber(result.TotalMaterialCost, 2) & vbCrLf
    End If
    UIQuoteReport = text & vbCrLf & _
        "Profile areas use supplied rectangle/trapezoid rules; holes, cutouts and reliefs are NOT deducted." & vbCrLf & _
        "Weight = profile area * thickness * density; material cost = weight * entered $/lb." & vbCrLf & _
        "Costs exclude purchased-sheet offcuts, cutting, bending, welding, finishing, labor and margin." & vbCrLf & _
        "Verify against known designs. This is not a final manufacturing quote." & vbCrLf & _
        "Stock sheet: 120 x 60 in. Sheet count and purchased-sheet cost NOT CALCULATED." & vbCrLf & _
        "Next inputs needed: cutting/nesting rules and cutout treatment."
End Function

Public Sub UIShowQuote(ByVal result As UIQuoteResult)
    Dim report As UserForm_QuoteReport
    Set report = New UserForm_QuoteReport
    report.SetReport UIQuoteReport(result)
    report.Show vbModal
    Unload report
End Sub
