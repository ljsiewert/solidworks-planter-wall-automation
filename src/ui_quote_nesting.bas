Attribute VB_Name = "ui_quote_nesting"
Option Explicit

Public Sub UINestQuote(ByVal result As UIQuoteResult, ByVal gap As Double, ByVal margin As Double)
    Dim pairedItems As Variant
    Dim singleItems As Variant
    Dim mixedItems As Variant
    Dim gussets() As String
    Dim part As UIQuoteComponent
    Dim nonGussetCount As Long
    Dim totalGussets As Long
    Dim i As Long
    Dim j As Long
    Dim pairShift As Double
    Dim pairWidth As Double
    Dim pairFits As Boolean
    Dim trial As Variant
    Dim sheets As Long
    Dim best As Long
    Dim mode As Long
    Dim usableLength As Double
    Dim usableWidth As Double
    Dim sortOrder As Long
    Dim packMode As Long
    Dim baselineSheets As Long
    Dim bestMethod As String
    Dim bestPlacements As Variant
    Dim elimPlacements As Variant
    Dim eliminatedAny As Boolean

    If gap < 0 Or margin < 0 Or margin >= 30 Then
        Err.Raise vbObjectError + 2450, "UINestQuote", "Spacing must be nonnegative and sheet-edge margin must be less than 30 inches."
    End If
    result.NestAvailable = False
    result.SheetCount = 0
    result.AreaMinimumSheets = 0
    result.UtilizationPercent = 0
    result.UnusedPercent = 0
    result.SheetPlacements = Empty
    result.PartGap = gap
    result.EdgeMargin = margin
    result.NestMethod = ""
    result.NestTrialCount = 0
    result.NestBaselineSheets = 0
    If Not result.HasValidTotals Then
        result.NestStatus = "Sheet estimate withheld: planter counts disagree."
        Exit Sub
    End If
    For i = 0 To UBound(result.Components)
        Set part = result.Components(i)
        If part.Quantity < 0 Or part.Quantity <> Int(part.Quantity) Then
            Err.Raise vbObjectError + 2450, "UINestQuote", "Nesting requires whole, nonnegative component quantities."
        End If
        If part.Quantity > 2000 Then
            result.NestStatus = "Sheet estimate unavailable: supported quote size is 1 to 2000 physical components."
            Exit Sub
        End If
        If part.TopWidth > 0 Then
            totalGussets = totalGussets + part.Quantity
        Else
            nonGussetCount = nonGussetCount + part.Quantity
        End If
    Next i
    If nonGussetCount + totalGussets = 0 Or nonGussetCount + totalGussets > 2000 Then
        result.NestStatus = "Sheet estimate unavailable: supported quote size is 1 to 2000 physical components."
        Exit Sub
    End If
    usableLength = 120 - 2 * margin
    usableWidth = 60 - 2 * margin
    If totalGussets > 0 Then
        ReDim gussets(0 To totalGussets - 1)
        j = 0
        For i = 0 To UBound(result.Components)
            Set part = result.Components(i)
            If part.TopWidth > 0 Then
                If Abs(part.Length - result.GussetHeight) > 0.000001 Or _
                   Abs(part.Width - result.GussetBottomWidth) > 0.000001 Or _
                   Abs(part.TopWidth - result.GussetTopWidth) > 0.000001 Then
                    Err.Raise vbObjectError + 2450, "UINestQuote", "Gusset pairing requires matching profiles."
                End If
                For mode = 1 To part.Quantity
                    gussets(j) = part.Name
                    j = j + 1
                Next mode
            End If
        Next i
        If Not ((result.GussetBottomWidth <= usableLength And result.GussetHeight <= usableWidth) Or _
                (result.GussetHeight <= usableLength And result.GussetBottomWidth <= usableWidth)) Then
            result.NestStatus = "Sheet estimate unavailable: " & gussets(0) & _
                " does not fit the 120 x 60 sheet with the selected margins."
            Exit Sub
        End If
        pairShift = gap * Sqr(1 + ((result.GussetBottomWidth - result.GussetTopWidth) / result.GussetHeight) ^ 2)
        pairWidth = result.GussetBottomWidth + result.GussetTopWidth + pairShift
        pairFits = ((pairWidth <= usableLength And result.GussetHeight <= usableWidth) Or _
                    (result.GussetHeight <= usableLength And pairWidth <= usableWidth))
    Else
        pairFits = False
    End If

    For i = 0 To UBound(result.Components)
        Set part = result.Components(i)
        If part.TopWidth = 0 And part.Quantity > 0 Then
            If Not ((part.Length <= usableLength And part.Width <= usableWidth) Or _
                    (part.Width <= usableLength And part.Length <= usableWidth)) Then
                result.NestStatus = "Sheet estimate unavailable: " & part.Name & _
                    " does not fit the 120 x 60 sheet with the selected margins."
                Exit Sub
            End If
        End If
    Next i

    best = nonGussetCount + totalGussets + 1
    pairedItems = UINestBuildItems(result, 0, pairFits, pairWidth, gussets, nonGussetCount, totalGussets)

    For mode = 0 To 27
        ' Keep the original four trials first so expanded search cannot increase sheet usage.
        If mode < 4 Then
            sortOrder = mode Mod 2
            packMode = mode \ 2
        Else
            sortOrder = (mode - 4) Mod 4
            packMode = (mode - 4) \ 4
        End If
        If mode < 4 Or packMode >= 2 Or sortOrder >= 2 Then
            UINestSort pairedItems, sortOrder
            trial = UINestPack(pairedItems, usableLength, usableWidth, gap, margin, packMode, sheets)
            If sheets < best Then
                best = sheets
                bestPlacements = trial
                bestMethod = "Guillotine - paired gussets"
            End If
        End If
    Next mode

    For mode = 0 To 7
        UINestSort pairedItems, mode Mod 4
        trial = UINestMaxRects(pairedItems, usableLength, usableWidth, gap, margin, mode \ 4, sheets)
        If sheets < best Then
            best = sheets
            bestPlacements = trial
            bestMethod = "MaxRects - paired gussets"
        End If
    Next mode

    baselineSheets = best
    result.NestBaselineSheets = baselineSheets

    Dim areaMinSheets As Long
    areaMinSheets = -Int(-result.TotalArea / 7200)

    If best > areaMinSheets And pairFits And totalGussets > 0 Then
        singleItems = UINestBuildItems(result, 1, pairFits, pairWidth, gussets, nonGussetCount, totalGussets)
        For mode = 0 To 7
            UINestSort singleItems, mode Mod 4
            trial = UINestMaxRects(singleItems, usableLength, usableWidth, gap, margin, mode \ 4, sheets)
            If sheets < best Then
                best = sheets
                bestPlacements = trial
                bestMethod = "MaxRects - single gussets"
            End If
        Next mode

        mixedItems = UINestBuildItems(result, 2, pairFits, pairWidth, gussets, nonGussetCount, totalGussets)
        For mode = 0 To 7
            UINestSort mixedItems, mode Mod 4
            trial = UINestMaxRects(mixedItems, usableLength, usableWidth, gap, margin, mode \ 4, sheets)
            If sheets < best Then
                best = sheets
                bestPlacements = trial
                bestMethod = "MaxRects - mixed gussets"
            End If
        Next mode
    End If

    result.NestTrialCount = 48

    eliminatedAny = False
    Do While best > 1
        If best <= areaMinSheets Then Exit Do
        If result.TotalArea > (best - 1) * usableLength * usableWidth Then Exit Do
        If UINestEliminateSheet(bestPlacements, best, usableLength, usableWidth, gap, margin, elimPlacements) Then
            best = best - 1
            bestPlacements = elimPlacements
            eliminatedAny = True
        Else
            Exit Do
        End If
    Loop
    If eliminatedAny Then
        bestMethod = bestMethod & " + sheet elimination"
    End If

    result.SheetCount = best
    result.SheetPlacements = bestPlacements
    result.NestMethod = bestMethod
    result.AreaMinimumSheets = -Int(-result.TotalArea / 7200)
    result.UtilizationPercent = 100 * result.TotalArea / (best * 7200)
    result.UnusedPercent = 100 - result.UtilizationPercent
    result.NestAvailable = True
    result.NestStatus = "Best of 48 estimating layouts (24 split + 24 MaxRects); Operations review required."
    If best < baselineSheets Then
        result.NestStatus = result.NestStatus & " Heuristic improvements reduced " & baselineSheets & " sheets to " & best & "."
    End If
    If best <= areaMinSheets Then
        result.NestStatus = result.NestStatus & " Area lower bound reached; remaining trials skipped as they cannot use fewer sheets."
    End If
End Sub

Private Function UINestBuildItems(ByVal result As UIQuoteResult, ByVal pairPolicy As Long, _
    ByVal pairFits As Boolean, ByVal pairWidth As Double, ByRef gussets() As String, _
    ByVal nonGussetCount As Long, ByVal totalGussets As Long) As Variant
    Dim items As Variant
    Dim pairsToMake As Long
    Dim singlesToMake As Long
    Dim totalItems As Long
    Dim itemIndex As Long
    Dim i As Long
    Dim j As Long
    Dim part As UIQuoteComponent
    Dim singleArea As Double
    Dim pairsMade As Long

    If totalGussets > 0 Then
        singleArea = (result.GussetTopWidth + result.GussetBottomWidth) / 2 * result.GussetHeight
    End If

    If Not pairFits Or totalGussets < 2 Then
        pairsToMake = 0
    ElseIf pairPolicy = 0 Then
        pairsToMake = totalGussets \ 2
    ElseIf pairPolicy = 1 Then
        pairsToMake = 0
    ElseIf pairPolicy = 2 Then
        pairsToMake = (totalGussets \ 2) \ 2
    End If
    singlesToMake = totalGussets - 2 * pairsToMake
    totalItems = nonGussetCount + pairsToMake + singlesToMake

    ReDim items(0 To totalItems - 1, 0 To 4)
    itemIndex = 0

    For i = 0 To UBound(result.Components)
        Set part = result.Components(i)
        If part.TopWidth = 0 Then
            For j = 1 To part.Quantity
                items(itemIndex, 0) = part.Name
                items(itemIndex, 1) = part.Length
                items(itemIndex, 2) = part.Width
                items(itemIndex, 3) = part.Area
                items(itemIndex, 4) = 1
                itemIndex = itemIndex + 1
            Next j
        End If
    Next i

    i = 0
    pairsMade = 0
    Do While i < totalGussets
        If pairsMade < pairsToMake And i + 1 < totalGussets Then
            items(itemIndex, 0) = gussets(i) & " + " & gussets(i + 1) & " (opposed pair)"
            items(itemIndex, 1) = pairWidth
            items(itemIndex, 2) = result.GussetHeight
            items(itemIndex, 3) = singleArea * 2
            items(itemIndex, 4) = 2
            pairsMade = pairsMade + 1
            i = i + 2
        Else
            items(itemIndex, 0) = gussets(i)
            items(itemIndex, 1) = result.GussetBottomWidth
            items(itemIndex, 2) = result.GussetHeight
            items(itemIndex, 3) = singleArea
            items(itemIndex, 4) = 1
            i = i + 1
        End If
        itemIndex = itemIndex + 1
    Loop

    UINestBuildItems = items
End Function

Private Function UINestEliminateSheet(ByVal placements As Variant, ByVal sheetCount As Long, _
    ByVal sheetLength As Double, ByVal sheetWidth As Double, ByVal gap As Double, _
    ByVal margin As Double, ByRef newPlacements As Variant) As Boolean
    Dim sheetArea() As Double
    Dim i As Long
    Dim j As Long
    Dim s As Long
    Dim minArea As Double
    Dim targetSheet As Long
    Dim relocCount As Long
    Dim bundles() As Variant
    Dim origIdx As Long
    Dim origW As Double
    Dim origH As Double
    Dim capacity As Long
    Dim freeCount As Long
    Dim free As Variant
    Dim px As Double
    Dim py As Double
    Dim pw As Double
    Dim ph As Double
    Dim chosen As Long
    Dim chosenRotation As Long
    Dim bestScore As Double
    Dim bestSecondary As Double
    Dim score As Double
    Dim secondary As Double
    Dim w0 As Double
    Dim h0 As Double
    Dim w90 As Double
    Dim h90 As Double
    Dim destSheet As Long
    Dim destX As Double
    Dim destY As Double
    Dim fitW As Double
    Dim fitH As Double
    Dim dx As Double
    Dim dy As Double
    Dim k As Long
    Dim swap As Variant

    UINestEliminateSheet = False
    If sheetCount <= 1 Then Exit Function

    ReDim sheetArea(0 To sheetCount)
    For i = 0 To UBound(placements)
        s = placements(i, 0)
        If s >= 1 And s <= sheetCount Then
            sheetArea(s) = sheetArea(s) + placements(i, 7)
        End If
    Next i

    minArea = 1E+100
    targetSheet = 0
    For s = sheetCount To 1 Step -1
        If sheetArea(s) < minArea Then
            minArea = sheetArea(s)
            targetSheet = s
        End If
    Next s
    If targetSheet = 0 Then Exit Function

    relocCount = 0
    For i = 0 To UBound(placements)
        If placements(i, 0) = targetSheet Then
            relocCount = relocCount + 1
        End If
    Next i
    If relocCount = 0 Then Exit Function

    ReDim bundles(0 To relocCount - 1, 0 To 5)
    relocCount = 0
    For i = 0 To UBound(placements)
        If placements(i, 0) = targetSheet Then
            bundles(relocCount, 0) = i
            If placements(i, 6) = 90 Then
                origW = placements(i, 5)
                origH = placements(i, 4)
            Else
                origW = placements(i, 4)
                origH = placements(i, 5)
            End If
            bundles(relocCount, 1) = origW
            bundles(relocCount, 2) = origH
            bundles(relocCount, 3) = placements(i, 7)
            bundles(relocCount, 4) = placements(i, 8)
            bundles(relocCount, 5) = placements(i, 1)
            relocCount = relocCount + 1
        End If
    Next i

    For i = 1 To relocCount - 1
        j = i
        Do While j > 0
            If bundles(j, 1) * bundles(j, 2) > bundles(j - 1, 1) * bundles(j - 1, 2) Then
                For k = 0 To 5
                    swap = bundles(j, k)
                    bundles(j, k) = bundles(j - 1, k)
                    bundles(j - 1, k) = swap
                Next k
                j = j - 1
            Else
                Exit Do
            End If
        Loop
    Next i

    capacity = 32
    freeCount = 0
    ReDim free(0 To 4, 0 To capacity - 1)
    For s = 1 To sheetCount
        If s <> targetSheet Then
            UINestFreeAdd free, freeCount, capacity, s, 0, 0, sheetLength + gap, sheetWidth + gap
            For i = 0 To UBound(placements)
                If placements(i, 0) = s Then
                    px = placements(i, 2) - margin
                    py = placements(i, 3) - margin
                    pw = placements(i, 4) + gap
                    ph = placements(i, 5) + gap
                    UINestFreeSubtract free, freeCount, capacity, s, px, py, pw, ph
                End If
            Next i
        End If
    Next s

    ReDim newPlacements(0 To UBound(placements), 0 To 8)
    For i = 0 To UBound(placements)
        For k = 0 To 8
            newPlacements(i, k) = placements(i, k)
        Next k
    Next i

    For i = 0 To relocCount - 1
        chosen = -1
        bestScore = 1E+100
        bestSecondary = 1E+100
        chosenRotation = 0
        w0 = bundles(i, 1) + gap
        h0 = bundles(i, 2) + gap
        w90 = bundles(i, 2) + gap
        h90 = bundles(i, 1) + gap
        For j = 0 To freeCount - 1
            If w0 <= free(3, j) And h0 <= free(4, j) Then
                score = free(3, j) * free(4, j) - w0 * h0
                secondary = free(3, j) - w0 + free(4, j) - h0 - UINestMax(free(3, j) - w0, free(4, j) - h0)
                If score < bestScore Or (score = bestScore And secondary < bestSecondary) Then
                    bestScore = score
                    bestSecondary = secondary
                    chosen = j
                    chosenRotation = 0
                End If
            End If
            If w90 <= free(3, j) And h90 <= free(4, j) Then
                score = free(3, j) * free(4, j) - w90 * h90
                secondary = free(3, j) - w90 + free(4, j) - h90 - UINestMax(free(3, j) - w90, free(4, j) - h90)
                If score < bestScore Or (score = bestScore And secondary < bestSecondary) Then
                    bestScore = score
                    bestSecondary = secondary
                    chosen = j
                    chosenRotation = 90
                End If
            End If
        Next j

        If chosen < 0 Then
            Exit Function
        End If

        destSheet = free(0, chosen)
        destX = free(1, chosen)
        destY = free(2, chosen)
        If chosenRotation = 0 Then
            fitW = w0
            fitH = h0
        Else
            fitW = w90
            fitH = h90
        End If

        UINestFreeSubtract free, freeCount, capacity, destSheet, destX, destY, fitW, fitH

        origIdx = bundles(i, 0)
        newPlacements(origIdx, 0) = destSheet
        newPlacements(origIdx, 2) = destX + margin
        newPlacements(origIdx, 3) = destY + margin
        newPlacements(origIdx, 4) = fitW - gap
        newPlacements(origIdx, 5) = fitH - gap
        newPlacements(origIdx, 6) = chosenRotation
    Next i

    For i = 0 To UBound(newPlacements)
        If newPlacements(i, 0) > targetSheet Then
            newPlacements(i, 0) = newPlacements(i, 0) - 1
        End If
    Next i

    For i = 0 To UBound(newPlacements)
        If newPlacements(i, 0) < 1 Or newPlacements(i, 0) > sheetCount - 1 Then Exit Function
        If newPlacements(i, 2) < margin - 0.000001 Or newPlacements(i, 3) < margin - 0.000001 Then Exit Function
        If newPlacements(i, 2) + newPlacements(i, 4) > 120 - margin + 0.000001 Then Exit Function
        If newPlacements(i, 3) + newPlacements(i, 5) > 60 - margin + 0.000001 Then Exit Function
        If newPlacements(i, 6) <> 0 And newPlacements(i, 6) <> 90 Then Exit Function
        For j = i + 1 To UBound(newPlacements)
            If newPlacements(i, 0) = newPlacements(j, 0) Then
                dx = UINestMax(newPlacements(j, 2) - newPlacements(i, 2) - newPlacements(i, 4), _
                              newPlacements(i, 2) - newPlacements(j, 2) - newPlacements(j, 4))
                dy = UINestMax(newPlacements(j, 3) - newPlacements(i, 3) - newPlacements(i, 5), _
                              newPlacements(i, 3) - newPlacements(j, 3) - newPlacements(j, 5))
                If dx < gap - 0.000001 And dy < gap - 0.000001 Then Exit Function
            End If
        Next j
    Next i

    UINestEliminateSheet = True
End Function

Private Sub UINestFreeSubtract(ByRef free As Variant, ByRef count As Long, ByRef capacity As Long, _
    ByVal sheet As Long, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double)
    Dim oldCount As Long
    Dim j As Long
    Dim fx As Double
    Dim fy As Double
    Dim fw As Double
    Dim fh As Double
    oldCount = count
    For j = 0 To oldCount - 1
        If free(0, j) = sheet Then
            fw = free(3, j)
            fh = free(4, j)
            fx = free(1, j)
            fy = free(2, j)
            If x < fx + fw And x + w > fx And y < fy + fh And y + h > fy Then
                If x > fx Then UINestFreeAdd free, count, capacity, sheet, fx, fy, x - fx, fh
                If x + w < fx + fw Then UINestFreeAdd free, count, capacity, sheet, x + w, fy, fx + fw - x - w, fh
                If y > fy Then UINestFreeAdd free, count, capacity, sheet, fx, fy, fw, y - fy
                If y + h < fy + fh Then UINestFreeAdd free, count, capacity, sheet, fx, y + h, fw, fy + fh - y - h
                free(3, j) = 0
            End If
        End If
    Next j
    UINestFreePrune free, count
End Sub

Private Function UINestMaxRects(ByVal items As Variant, ByVal sheetLength As Double, _
    ByVal sheetWidth As Double, ByVal gap As Double, ByVal margin As Double, _
    ByVal fitMode As Long, ByRef sheetCount As Long) As Variant
    Dim free As Variant
    Dim placed() As Variant
    Dim count As Long
    Dim capacity As Long
    Dim i As Long
    Dim j As Long
    Dim r As Long
    Dim chosen As Long
    Dim rotation As Long
    Dim w As Double
    Dim h As Double
    Dim score As Double
    Dim secondary As Double
    Dim bestScore As Double
    Dim bestSecondary As Double
    Dim x As Double
    Dim y As Double
    Dim sheet As Long
    capacity = 32
    ReDim free(0 To 4, 0 To capacity - 1)
    ReDim placed(0 To UBound(items), 0 To 8)
    sheetCount = 0
    For i = 0 To UBound(items)
        chosen = -1
        bestScore = 1E+100
        bestSecondary = 1E+100
        For j = 0 To count - 1
            For r = 0 To 1
                w = items(i, 1 + r) + gap
                h = items(i, 2 - r) + gap
                If w <= free(3, j) And h <= free(4, j) Then
                    score = free(3, j) * free(4, j) - w * h
                    secondary = free(3, j) - w + free(4, j) - h - UINestMax(free(3, j) - w, free(4, j) - h)
                    If fitMode = 1 Then
                        score = secondary
                        secondary = UINestMax(free(3, j) - w, free(4, j) - h)
                    End If
                    If score < bestScore Or (score = bestScore And secondary < bestSecondary) Then
                        bestScore = score
                        bestSecondary = secondary
                        chosen = j
                        rotation = r
                    End If
                End If
            Next r
        Next j
        If chosen < 0 Then
            sheetCount = sheetCount + 1
            chosen = count
            UINestFreeAdd free, count, capacity, sheetCount, 0, 0, sheetLength + gap, sheetWidth + gap
            rotation = 0
            If items(i, 1) > sheetLength Or items(i, 2) > sheetWidth Then rotation = 1
        End If
        sheet = free(0, chosen)
        x = free(1, chosen)
        y = free(2, chosen)
        w = items(i, 1 + rotation) + gap
        h = items(i, 2 - rotation) + gap
        placed(i, 0) = sheet
        placed(i, 1) = items(i, 0)
        placed(i, 2) = x + margin
        placed(i, 3) = y + margin
        placed(i, 4) = w - gap
        placed(i, 5) = h - gap
        placed(i, 6) = rotation * 90
        placed(i, 7) = items(i, 3)
        placed(i, 8) = items(i, 4)
        UINestFreeSubtract free, count, capacity, sheet, x, y, w, h
    Next i
    UINestMaxRects = placed
End Function

Private Sub UINestFreeAdd(ByRef free As Variant, ByRef count As Long, ByRef capacity As Long, _
    ByVal sheet As Long, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double)
    If count = capacity Then
        capacity = capacity * 2
        ReDim Preserve free(0 To 4, 0 To capacity - 1)
    End If
    free(0, count) = sheet
    free(1, count) = x
    free(2, count) = y
    free(3, count) = w
    free(4, count) = h
    count = count + 1
End Sub

Private Sub UINestFreePrune(ByRef free As Variant, ByRef count As Long)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim kept As Long
    Dim contained As Boolean
    For i = 0 To count - 1
        If free(3, i) > 0 And free(4, i) > 0 Then
            For k = 0 To 4
                free(k, kept) = free(k, i)
            Next k
            kept = kept + 1
        End If
    Next i
    count = kept
    i = 0
    Do While i < count
        contained = False
        For j = 0 To count - 1
            If i <> j And free(0, i) = free(0, j) Then
                If free(1, i) >= free(1, j) And free(2, i) >= free(2, j) And _
                    free(1, i) + free(3, i) <= free(1, j) + free(3, j) And _
                    free(2, i) + free(4, i) <= free(2, j) + free(4, j) Then
                    contained = True
                    Exit For
                End If
            End If
        Next j
        If contained Then
            count = count - 1
            For k = 0 To 4
                free(k, i) = free(k, count)
            Next k
        Else
            i = i + 1
        End If
    Loop
End Sub

Private Sub UINestSort(ByRef items As Variant, ByVal order As Long)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim a As Double
    Dim b As Double
    Dim swap As Variant
    For i = 1 To UBound(items)
        j = i
        Do While j > 0
            a = items(j - 1, 1) * items(j - 1, 2)
            b = items(j, 1) * items(j, 2)
            If order = 1 Then
                a = UINestMax(items(j - 1, 1), items(j - 1, 2))
                b = UINestMax(items(j, 1), items(j, 2))
            ElseIf order = 2 Then
                a = items(j - 1, 1) + items(j - 1, 2) - UINestMax(items(j - 1, 1), items(j - 1, 2))
                b = items(j, 1) + items(j, 2) - UINestMax(items(j, 1), items(j, 2))
            ElseIf order = 3 Then
                a = items(j - 1, 1) + items(j - 1, 2)
                b = items(j, 1) + items(j, 2)
            End If
            If a >= b Then Exit Do
            For k = 0 To 4
                swap = items(j - 1, k)
                items(j - 1, k) = items(j, k)
                items(j, k) = swap
            Next k
            j = j - 1
        Loop
    Next i
End Sub

Private Function UINestMax(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then UINestMax = a Else UINestMax = b
End Function

Private Function UINestPack(ByVal items As Variant, ByVal sheetLength As Double, _
    ByVal sheetWidth As Double, ByVal gap As Double, ByVal margin As Double, _
    ByVal splitMode As Long, ByRef sheetCount As Long) As Variant
    Dim free() As Double
    Dim placed() As Variant
    Dim count As Long
    Dim i As Long
    Dim j As Long
    Dim r As Long
    Dim chosen As Long
    Dim chosenRotation As Long
    Dim w As Double
    Dim h As Double
    Dim fitW As Double
    Dim fitH As Double
    Dim score As Double
    Dim bestScore As Double
    Dim x As Double
    Dim y As Double
    Dim fw As Double
    Dim fh As Double
    Dim sheet As Double
    Dim bestSecondary As Double
    Dim secondary As Double
    Dim split As Long
    ReDim free(0 To 2 * (UBound(items) + 1), 0 To 4)
    ReDim placed(0 To UBound(items), 0 To 8)
    sheetCount = 0
    For i = 0 To UBound(items)
        chosen = -1
        bestScore = 1E+100
        bestSecondary = 1E+100
        For j = 0 To count - 1
            For r = 0 To 1
                w = items(i, 1 + r) + gap
                h = items(i, 2 - r) + gap
                If w <= free(j, 3) And h <= free(j, 4) Then
                    score = free(j, 3) * free(j, 4) - w * h
                    secondary = 0
                    If splitMode >= 3 Then
                        score = free(j, 3) - w + free(j, 4) - h - UINestMax(free(j, 3) - w, free(j, 4) - h)
                        secondary = UINestMax(free(j, 3) - w, free(j, 4) - h)
                    End If
                    If score < bestScore Or (score = bestScore And secondary < bestSecondary) Then
                        bestScore = score
                        bestSecondary = secondary
                        chosen = j
                        chosenRotation = r
                    End If
                End If
            Next r
        Next j
        If chosen < 0 Then
            sheetCount = sheetCount + 1
            chosen = count
            free(count, 0) = sheetCount
            free(count, 1) = 0
            free(count, 2) = 0
            free(count, 3) = sheetLength + gap
            free(count, 4) = sheetWidth + gap
            count = count + 1
            chosenRotation = 0
            If items(i, 1) > sheetLength Or items(i, 2) > sheetWidth Then chosenRotation = 1
        End If
        fitW = items(i, 1 + chosenRotation) + gap
        fitH = items(i, 2 - chosenRotation) + gap
        sheet = free(chosen, 0)
        x = free(chosen, 1)
        y = free(chosen, 2)
        fw = free(chosen, 3)
        fh = free(chosen, 4)
        placed(i, 0) = sheet
        placed(i, 1) = items(i, 0)
        placed(i, 2) = x + margin
        placed(i, 3) = y + margin
        placed(i, 4) = fitW - gap
        placed(i, 5) = fitH - gap
        placed(i, 6) = chosenRotation * 90
        placed(i, 7) = items(i, 3)
        placed(i, 8) = items(i, 4)
        free(chosen, 0) = sheet
        free(count, 0) = sheet
        free(chosen, 1) = x + fitW
        free(chosen, 2) = y
        free(chosen, 3) = fw - fitW
        free(count, 1) = x
        free(count, 2) = y + fitH
        free(count, 4) = fh - fitH
        split = splitMode Mod 3
        If split = 2 Then
            If UINestMax((fw - fitW) * fitH, fw * (fh - fitH)) >= _
                UINestMax((fw - fitW) * fh, fitW * (fh - fitH)) Then
                split = 0
            Else
                split = 1
            End If
        End If
        If split = 0 Then
            free(chosen, 4) = fitH
            free(count, 3) = fw
        Else
            free(chosen, 4) = fh
            free(count, 3) = fitW
        End If
        count = count + 1
    Next i
    UINestPack = placed
End Function

Public Function UINestReport(ByVal result As UIQuoteResult) As String
    Dim text As String
    Dim sheet As Long
    Dim i As Long
    Dim area As Double
    Dim quantity As Long
    text = vbCrLf & "SHEET LAYOUT ESTIMATE - 120 x 60 in" & vbCrLf & result.NestStatus & vbCrLf & _
        "Part spacing: " & UIInchesText(result.PartGap) & " in; edge margin: " & UIInchesText(result.EdgeMargin) & " in." & vbCrLf & _
        "Provisional 0.25-inch defaults are estimating assumptions, NOT a mandatory laser standard." & vbCrLf
    If Not result.NestAvailable Then
        UINestReport = text
        Exit Function
    End If
    text = text & "Estimated sheets: " & result.SheetCount & "; area-only lower bound: " & result.AreaMinimumSheets & vbCrLf & _
        "Profile utilization: " & FormatNumber(result.UtilizationPercent, 2) & "%; unused/scrap: " & FormatNumber(result.UnusedPercent, 2) & "%" & vbCrLf & _
        "Unused area includes reusable remnants, margins, gaps and shape offcuts; not necessarily discarded scrap." & vbCrLf & _
        "No common-line cutting, lead-in paths, clamps, heat sequencing or remnant reuse is modeled." & vbCrLf
    text = text & "Gussets pair at 180 degrees; pair envelopes use 0/90 rotations." & vbCrLf & _
        "Best of " & result.NestTrialCount & " layouts: 24 guillotine splits and " & (result.NestTrialCount - 24) & " MaxRects trials; not a guaranteed minimum." & vbCrLf
    For sheet = 1 To result.SheetCount
        area = 0
        quantity = 0
        For i = 0 To UBound(result.SheetPlacements)
            If result.SheetPlacements(i, 0) = sheet Then
                area = area + result.SheetPlacements(i, 7)
                quantity = quantity + result.SheetPlacements(i, 8)
            End If
        Next i
        text = text & vbCrLf & "Sheet " & sheet & ": " & quantity & " parts; utilization " & FormatNumber(area / 72, 2) & "%" & vbCrLf
        For i = 0 To UBound(result.SheetPlacements)
            If result.SheetPlacements(i, 0) = sheet Then
                text = text & result.SheetPlacements(i, 1) & ": x=" & FormatNumber(result.SheetPlacements(i, 2), 3) & _
                    ", y=" & FormatNumber(result.SheetPlacements(i, 3), 3) & _
                    ", envelope=" & FormatNumber(result.SheetPlacements(i, 4), 3) & " x " & FormatNumber(result.SheetPlacements(i, 5), 3) & _
                    " in; rotation=" & result.SheetPlacements(i, 6) & " deg" & vbCrLf
            End If
        Next i
    Next sheet
    UINestReport = text
End Function
