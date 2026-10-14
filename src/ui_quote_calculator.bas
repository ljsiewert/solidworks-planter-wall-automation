Attribute VB_Name = "ui_quote_calculator"
Option Explicit

Public Function UICalculateQuote(ByVal walls As Long, ByVal orientation As String, _
    ByVal returnLabel As String, ByVal material As String, ByVal thicknessLabel As String, _
    ByVal length As Double, ByVal width As Double, ByVal height As Double, _
    Optional ByVal pricePerPound As Double = 0.7, _
    Optional ByVal partGap As Double = 0.25, Optional ByVal edgeMargin As Double = 0.25) As UIQuoteResult
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
    UINestQuote result, partGap, edgeMargin
    UIQuotePricing result, False, UIQuoteDefaultRate("Cut"), UIQuoteDefaultRate("Bend"), _
        UIQuoteDefaultRate("Weld"), UIQuoteDefaultRate("Paint"), UIQuoteDefaultRate("Margin")
    Set UICalculateQuote = result
End Function

' Standard placeholder rates until time studies are available; editable under Properties.
Public Function UIQuoteDefaultRate(ByVal name As String) As Double
    Select Case name
        Case "Cut": UIQuoteDefaultRate = 75
        Case "Bend": UIQuoteDefaultRate = 5
        Case "Weld": UIQuoteDefaultRate = 8
        Case "Paint": UIQuoteDefaultRate = 2.5
        Case "Margin": UIQuoteDefaultRate = 30
        Case Else: Err.Raise vbObjectError + 2400, "UIQuoteDefaultRate", "Unknown rate: " & name
    End Select
End Function

' Labor = cutting per purchased sheet + bending per bent part (returns only) + welding per part.
' Paint covers both sides of every profile. Material uses full purchased sheets so scrap is paid for.
Public Sub UIQuotePricing(ByVal result As UIQuoteResult, ByVal painted As Boolean, _
    ByVal cutRate As Double, ByVal bendRate As Double, ByVal weldRate As Double, _
    ByVal paintRate As Double, ByVal marginPercent As Double)
    Dim i As Long
    If cutRate < 0 Or bendRate < 0 Or weldRate < 0 Or paintRate < 0 Then
        Err.Raise vbObjectError + 2400, "UIQuotePricing", "Labor and paint rates cannot be negative."
    End If
    result.Painted = painted
    result.CutRate = cutRate
    result.BendRate = bendRate
    result.WeldRate = weldRate
    result.PaintRate = paintRate
    result.HasPricing = result.HasValidTotals And result.NestAvailable And result.SheetCount > 0
    result.BentParts = 0
    result.WeldedParts = 0
    For i = 0 To UBound(result.Components)
        result.WeldedParts = result.WeldedParts + result.Components(i).Quantity
    Next i
    If result.ReturnLabel <> "No return" Then
        result.BentParts = result.Components(0).Quantity + result.Components(2).Quantity + _
            result.Components(4).Quantity + result.Components(5).Quantity
    End If
    result.SheetWeight = 0
    result.SheetMaterialCost = 0
    result.CutCost = 0
    If result.HasPricing Then
        result.SheetWeight = result.SheetCount * 7200 * result.Thickness * result.Density
        result.SheetMaterialCost = result.SheetWeight * result.PricePerPound
        result.CutCost = cutRate * result.SheetCount
    End If
    result.BendCost = bendRate * result.BentParts
    result.WeldCost = weldRate * result.WeldedParts
    result.LaborCost = result.CutCost + result.BendCost + result.WeldCost
    result.PaintArea = 0
    If painted Then result.PaintArea = 2 * result.TotalArea / 144
    result.PaintCost = paintRate * result.PaintArea
    result.TotalCost = result.SheetMaterialCost + result.LaborCost + result.PaintCost
    UIQuoteSetMargin result, marginPercent
End Sub

' Gross margin: profit is the given percentage of the sell price.
Public Sub UIQuoteSetMargin(ByVal result As UIQuoteResult, ByVal marginPercent As Double)
    If marginPercent < 0 Or marginPercent > 95 Then
        Err.Raise vbObjectError + 2400, "UIQuoteSetMargin", "Margin must be between 0% and 95%."
    End If
    result.MarginPercent = marginPercent
    result.SellPrice = result.TotalCost / (1 - marginPercent / 100)
    result.Profit = result.SellPrice - result.TotalCost
End Sub

Public Function UIQuoteParseRate(ByVal text As String, ByVal label As String) As Double
    Dim cleaned As String
    cleaned = Trim$(text)
    If Len(cleaned) > 0 And InStr(cleaned, "0") > 0 And Len(Replace(Replace(cleaned, "0", ""), ".", "")) = 0 _
        And Len(cleaned) - Len(Replace(cleaned, ".", "")) <= 1 Then
        UIQuoteParseRate = 0
        Exit Function
    End If
    On Error GoTo InvalidRate
    UIQuoteParseRate = UIParseInches(cleaned)
    Exit Function
InvalidRate:
    Err.Raise vbObjectError + 2400, "UIQuoteParseRate", _
        label & ": enter 0 or a positive decimal number (no $, % or commas)."
End Function

' Rows: caption, value, tooltip breakdown for material, labor, paint, total cost and profit.
Public Function UIQuotePricingCards(ByVal result As UIQuoteResult) As Variant
    Dim cards(0 To 4, 0 To 2) As Variant
    cards(0, 0) = "Material (purchased sheets)"
    cards(1, 0) = "Labor (cut / bend / weld)"
    cards(2, 0) = "Paint"
    cards(3, 0) = "Total cost"
    cards(4, 0) = "Margin %  |  profit"
    If Not result.HasPricing Then
        cards(0, 1) = "Withheld"
        cards(1, 1) = "Withheld"
        cards(2, 1) = "Withheld"
        cards(3, 1) = "Withheld"
        cards(4, 1) = "Withheld"
        cards(0, 2) = "Pricing needs valid quantities and a sheet layout."
        cards(1, 2) = cards(0, 2)
        cards(2, 2) = cards(0, 2)
        cards(3, 2) = cards(0, 2)
        cards(4, 2) = cards(0, 2)
        UIQuotePricingCards = cards
        Exit Function
    End If
    cards(0, 1) = "$" & FormatNumber(result.SheetMaterialCost, 2)
    cards(0, 2) = result.SheetCount & " sheets x " & FormatNumber(result.SheetWeight / result.SheetCount, 2) & _
        " lb x $" & FormatNumber(result.PricePerPound, 2) & "/lb (parts only: $" & FormatNumber(result.TotalMaterialCost, 2) & ")"
    cards(1, 1) = "$" & FormatNumber(result.LaborCost, 2)
    cards(1, 2) = "Cut " & result.SheetCount & " sheets x $" & FormatNumber(result.CutRate, 2) & _
        " + bend " & result.BentParts & " parts x $" & FormatNumber(result.BendRate, 2) & _
        " + weld " & result.WeldedParts & " parts x $" & FormatNumber(result.WeldRate, 2)
    If result.Painted Then
        cards(2, 1) = "$" & FormatNumber(result.PaintCost, 2)
        cards(2, 2) = FormatNumber(result.PaintArea, 2) & " ft^2 (both sides) x $" & FormatNumber(result.PaintRate, 2) & "/ft^2"
    Else
        cards(2, 1) = "Not painted"
        cards(2, 2) = "Turn on Painted next to Material to include paint."
    End If
    cards(3, 1) = "$" & FormatNumber(result.TotalCost, 2)
    cards(3, 2) = "Material + labor + paint"
    cards(4, 1) = "$" & FormatNumber(result.Profit, 2)
    cards(4, 2) = "Sell price = total cost / (1 - margin)"
    UIQuotePricingCards = cards
End Function

Public Function UIQuotePricingReport(ByVal result As UIQuoteResult) As String
    If Not result.HasPricing Then
        UIQuotePricingReport = "PRICING: WITHHELD - needs valid quantities and a sheet layout." & vbCrLf
        Exit Function
    End If
    UIQuotePricingReport = "PRICING (placeholder rates until time studies are available)" & vbCrLf & _
        "Material, purchased sheets: " & result.SheetCount & " x 120 x 60 in = " & _
        FormatNumber(result.SheetWeight, 2) & " lb x $" & FormatNumber(result.PricePerPound, 2) & "/lb = $" & _
        FormatNumber(result.SheetMaterialCost, 2) & vbCrLf & _
        "Labor - cutting: " & result.SheetCount & " sheets x $" & FormatNumber(result.CutRate, 2) & " = $" & FormatNumber(result.CutCost, 2) & vbCrLf & _
        "Labor - bending: " & result.BentParts & " bent parts x $" & FormatNumber(result.BendRate, 2) & " = $" & FormatNumber(result.BendCost, 2) & _
        IIf(result.BentParts = 0, " (no return - no bending)", "") & vbCrLf & _
        "Labor - welding: " & result.WeldedParts & " parts x $" & FormatNumber(result.WeldRate, 2) & " = $" & FormatNumber(result.WeldCost, 2) & vbCrLf & _
        "Labor total: $" & FormatNumber(result.LaborCost, 2) & vbCrLf & _
        "Paint: " & IIf(result.Painted, FormatNumber(result.PaintArea, 2) & " ft^2 (both sides) x $" & _
            FormatNumber(result.PaintRate, 2) & "/ft^2 = $" & FormatNumber(result.PaintCost, 2), "not painted") & vbCrLf & _
        "TOTAL COST: $" & FormatNumber(result.TotalCost, 2) & vbCrLf & _
        "Margin: " & FormatNumber(result.MarginPercent, 2) & "% of sell price; profit $" & FormatNumber(result.Profit, 2) & vbCrLf & _
        "SELL PRICE: $" & FormatNumber(result.SellPrice, 2) & vbCrLf
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
    Set parts(4) = UIQuotePart(result, "Left Side Wall", "Side ends", result.Height / 2, result.FlatHeight, 0, result.SideWallRuns, "H/2 x calibrated flat height; one per side run")
    Set parts(5) = UIQuotePart(result, "Right Side Wall", "Side ends", result.Height / 2, result.FlatHeight, 0, result.SideWallRuns, "H/2 x calibrated flat height; one per side run")
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
    UIQuoteReport = text & vbCrLf & UIQuotePricingReport(result) & vbCrLf & _
        "Profile areas use supplied rectangle/trapezoid rules; holes, cutouts and reliefs are NOT deducted." & vbCrLf & _
        "Weight = profile area * thickness * density; part material cost = weight * entered $/lb." & vbCrLf & _
        "Quoted material cost uses full purchased sheets, so offcuts and scrap are included." & vbCrLf & _
        "Labor and paint use placeholder rates; finishing other than paint, freight and hardware are excluded." & vbCrLf & _
        "Verify against known designs. This is not a final manufacturing quote." & vbCrLf & _
        "Next: Operations validation of spacing, margins, layout, labor rates and cutout treatment." & vbCrLf & _
        UINestReport(result)
End Function

' Rows: 0 = headline (caption, value), 1 = design details, 2-6 = metric tiles.
Public Function UIQuoteSummaryCards(ByVal result As UIQuoteResult) As Variant
    Dim cards(0 To 6, 0 To 1) As Variant
    Dim design As String
    design = result.Material & IIf(result.Painted, " (painted)", "") & "  |  " & IIf(result.Thickness = 0.25, "1/4", "3/16") & _
        " in  |  $" & FormatNumber(result.PricePerPound, 2) & "/lb" & vbCrLf & _
        CStr(result.WallCount) & " walls  |  " & result.Orientation & "  |  " & result.ReturnLabel & vbCrLf & _
        "L x W x H: " & FormatNumber(result.OverallLength, 3) & " x " & _
        FormatNumber(result.OverallWidth, 3) & " x " & FormatNumber(result.Height, 3) & " in"
    cards(1, 0) = "Design"
    cards(2, 0) = "Profile weight"
    cards(3, 0) = "Profile area"
    cards(4, 0) = "120 x 60 sheets"
    cards(5, 0) = "Sheet utilization"
    cards(6, 0) = "Unused / potential scrap"
    If result.HasValidTotals Then
        cards(0, 0) = "PROFILE MATERIAL ESTIMATE"
        cards(0, 1) = "$" & FormatNumber(result.TotalMaterialCost, 2)
        If result.HasPricing Then
            cards(0, 0) = "QUOTE SELL PRICE  (" & FormatNumber(result.MarginPercent, 2) & "% margin)"
            cards(0, 1) = "$" & FormatNumber(result.SellPrice, 2)
        End If
        cards(1, 1) = design
        cards(2, 1) = FormatNumber(result.TotalWeight, 2) & " lb"
        cards(3, 1) = FormatNumber(result.TotalArea / 144, 2) & " ft^2"
    Else
        cards(0, 0) = "REVIEW REQUIRED - TOTALS WITHHELD"
        cards(0, 1) = "Review required"
        cards(1, 1) = result.Warnings & "Quantities, total weight and costs are withheld."
        cards(2, 1) = "Withheld"
        cards(3, 1) = "Withheld"
    End If
    If result.NestAvailable Then
        cards(4, 1) = CStr(result.SheetCount)
        cards(5, 1) = FormatNumber(result.UtilizationPercent, 2) & "%"
        cards(6, 1) = FormatNumber(result.UnusedPercent, 2) & "%"
    Else
        cards(4, 1) = "N/A"
        cards(5, 1) = "N/A"
        cards(6, 1) = "N/A"
    End If
    UIQuoteSummaryCards = cards
End Function

Public Function UIQuoteComponentRows(ByVal result As UIQuoteResult) As Variant
    UIQuoteComponentRows = UIQuoteTableRows(result, False)
End Function

Public Function UIQuoteGussetRows(ByVal result As UIQuoteResult) As Variant
    UIQuoteGussetRows = UIQuoteTableRows(result, True)
End Function

Private Function UIQuoteTableRows(ByVal result As UIQuoteResult, ByVal gussets As Boolean) As Variant
    Dim rows() As Variant
    Dim part As UIQuoteComponent
    Dim i As Long
    Dim count As Long
    Dim row As Long
    Dim weightColumn As Long
    For i = 0 To UBound(result.Components)
        Set part = result.Components(i)
        If (part.TopWidth > 0) = gussets Then count = count + 1
    Next i
    If count = 0 Then
        Err.Raise vbObjectError + 2400, "UIQuoteTableRows", "No components found for the requested quote table."
    End If
    weightColumn = IIf(gussets, 5, 4)
    ' Zero-quantity parts are hidden; untrusted totals keep every row for review.
    count = 0
    For i = 0 To UBound(result.Components)
        Set part = result.Components(i)
        If (part.TopWidth > 0) = gussets And (Not result.HasValidTotals Or part.Quantity > 0) Then count = count + 1
    Next i
    If count = 0 Then
        ReDim rows(0 To 0, 0 To weightColumn + 1)
        rows(0, 0) = IIf(gussets, "No gussets required", "No components required")
        UIQuoteTableRows = rows
        Exit Function
    End If
    ReDim rows(0 To count - 1, 0 To weightColumn + 1)
    For i = 0 To UBound(result.Components)
        Set part = result.Components(i)
        If (part.TopWidth > 0) = gussets And (Not result.HasValidTotals Or part.Quantity > 0) Then
            rows(row, 0) = part.Name
            rows(row, 2) = FormatNumber(part.Length, 3)
            If gussets Then
                rows(row, 3) = FormatNumber(part.TopWidth, 3)
                rows(row, 4) = FormatNumber(part.Width, 3)
            Else
                rows(row, 3) = FormatNumber(part.Width, 3)
            End If
            If result.HasValidTotals Then
                rows(row, 1) = FormatNumber(part.Quantity, 0)
                rows(row, weightColumn) = FormatNumber(part.TotalWeight, 2)
                rows(row, weightColumn + 1) = "$" & FormatNumber(part.TotalMaterialCost, 2)
            Else
                rows(row, 1) = "Review"
                rows(row, weightColumn) = "Withheld"
                rows(row, weightColumn + 1) = "Withheld"
            End If
            row = row + 1
        End If
    Next i
    UIQuoteTableRows = rows
End Function

Public Function UIQuoteTotalRows(ByVal result As UIQuoteResult) As Variant
    Dim rows(0 To 0, 0 To 2) As Variant
    rows(0, 0) = "TOTAL MATERIAL - ALL COMPONENTS"
    If result.HasValidTotals Then
        rows(0, 1) = FormatNumber(result.TotalWeight, 2) & " lb"
        rows(0, 2) = "$" & FormatNumber(result.TotalMaterialCost, 2)
    Else
        rows(0, 1) = "Withheld"
        rows(0, 2) = "Withheld"
    End If
    UIQuoteTotalRows = rows
End Function

Public Sub UIShowQuote(ByVal result As UIQuoteResult)
    Dim report As UserForm_QuoteReport
    Set report = New UserForm_QuoteReport
    report.SetResult result
    report.Show vbModal
    Unload report
End Sub

Public Function UIQuoteSheetStats(ByVal result As UIQuoteResult) As Variant
    Dim rows() As Variant
    Dim i As Long
    Dim sheet As Long
    Dim area As Double
    If Not result.NestAvailable Or result.SheetCount < 1 Then
        Err.Raise vbObjectError + 2400, "UIQuoteSheetStats", "Sheet usage is unavailable: " & result.NestStatus
    End If
    ReDim rows(0 To result.SheetCount, 0 To 5)
    For sheet = 0 To result.SheetCount
        rows(sheet, 0) = sheet + 1
        rows(sheet, 1) = 0
        rows(sheet, 2) = 0#
    Next sheet
    For i = 0 To UBound(result.SheetPlacements)
        sheet = result.SheetPlacements(i, 0) - 1
        If sheet < 0 Or sheet >= result.SheetCount Then
            Err.Raise vbObjectError + 2400, "UIQuoteSheetStats", "Invalid sheet index in layout."
        End If
        rows(sheet, 1) = rows(sheet, 1) + result.SheetPlacements(i, 8)
        rows(sheet, 2) = rows(sheet, 2) + result.SheetPlacements(i, 7)
    Next i
    For sheet = 0 To result.SheetCount - 1
        area = rows(sheet, 2)
        If area < 0 Or area > 7200.000001 Then
            Err.Raise vbObjectError + 2400, "UIQuoteSheetStats", "Invalid sheet profile area."
        End If
        rows(result.SheetCount, 1) = rows(result.SheetCount, 1) + rows(sheet, 1)
        rows(result.SheetCount, 2) = rows(result.SheetCount, 2) + area
        rows(sheet, 2) = area / 144
        rows(sheet, 3) = (7200 - area) / 144
        rows(sheet, 4) = area / 72
        rows(sheet, 5) = 100 - rows(sheet, 4)
    Next sheet
    area = rows(result.SheetCount, 2)
    If Abs(area - result.TotalArea) > 0.000001 * result.SheetCount Then
        Err.Raise vbObjectError + 2400, "UIQuoteSheetStats", "Sheet areas disagree with component totals."
    End If
    rows(result.SheetCount, 0) = "TOTAL"
    rows(result.SheetCount, 2) = area / 144
    rows(result.SheetCount, 3) = (result.SheetCount * 7200 - area) / 144
    rows(result.SheetCount, 4) = area / (result.SheetCount * 72)
    rows(result.SheetCount, 5) = 100 - rows(result.SheetCount, 4)
    UIQuoteSheetStats = rows
End Function

Public Function UIQuoteSheetRows(ByVal result As UIQuoteResult) As Variant
    Dim rows As Variant
    Dim i As Long
    Dim j As Long
    If Not result.NestAvailable Then
        ReDim rows(0 To 0, 0 To 5)
        rows(0, 0) = "Unavailable"
        For j = 1 To 5
            rows(0, j) = "Withheld"
        Next j
    Else
        rows = UIQuoteSheetStats(result)
        For i = 0 To UBound(rows)
            rows(i, 1) = FormatNumber(rows(i, 1), 0)
            rows(i, 2) = FormatNumber(rows(i, 2), 2)
            rows(i, 3) = FormatNumber(rows(i, 3), 2)
            rows(i, 4) = FormatNumber(rows(i, 4), 2) & "%"
            rows(i, 5) = FormatNumber(rows(i, 5), 2) & "%"
        Next i
    End If
    UIQuoteSheetRows = rows
End Function
