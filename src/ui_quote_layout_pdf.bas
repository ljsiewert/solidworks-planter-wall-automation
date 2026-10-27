Attribute VB_Name = "ui_quote_layout_pdf"
Option Explicit

Public Sub UIPreviewQuoteLayoutPDF(ByVal result As UIQuoteResult)
    Dim fso As Object
    Dim shell As Object
    Dim path As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = fso.BuildPath(CStr(fso.GetSpecialFolder(2)), "Planter_Layout_" & fso.GetBaseName(fso.GetTempName) & ".pdf")
    UISaveQuoteLayoutPDF result, path
    Set shell = CreateObject("Shell.Application")
    shell.ShellExecute path, "", "", "open", 1
End Sub

Public Function UIQuoteLayoutPDF(ByVal result As UIQuoteResult) As String
    Dim text As String
    Dim kids As String
    Dim contents As String
    Dim offsets() As Long
    Dim sheet As Long
    Dim pageId As Long
    Dim count As Long
    Dim startXref As Long
    Dim i As Long
    Dim usagePages As Long
    Dim page As Long
    Dim pageCount As Long
    If Not result.NestAvailable Or result.SheetCount < 1 Then
        Err.Raise vbObjectError + 2460, "UIQuoteLayoutPDF", "No valid sheet layout is available to export."
    End If
    usagePages = -Int(-(result.SheetCount + 1) / 24)
    pageCount = result.SheetCount + 2 + usagePages
    count = 4 + 2 * pageCount
    ReDim offsets(0 To count)
    text = "%PDF-1.4" & vbCrLf
    For page = 0 To pageCount - 1
        kids = kids & CStr(5 + page * 2) & " 0 R "
    Next page
    text = text & UIPDFObject(1, "<< /Type /Catalog /Pages 2 0 R >>", Len(text), offsets)
    text = text & UIPDFObject(2, "<< /Type /Pages /Count " & pageCount & " /Kids [" & kids & "] >>", Len(text), offsets)
    text = text & UIPDFObject(3, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>", Len(text), offsets)
    text = text & UIPDFObject(4, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>", Len(text), offsets)
    For page = 0 To pageCount - 1
        pageId = 5 + page * 2
        text = text & UIPDFObject(pageId, "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 792 612] " & _
            "/Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> /Contents " & (pageId + 1) & " 0 R >>", Len(text), offsets)
        If page = 0 Then
            contents = UIPDFSummary(result)
        ElseIf page <= usagePages Then
            contents = UIPDFUsage(result, page - 1, usagePages)
        ElseIf page = usagePages + 1 Then
            contents = UIPDFParts(result)
        Else
            sheet = page - usagePages - 1
            contents = UIPDFSheet(result, sheet)
        End If
        text = text & UIPDFObject(pageId + 1, "<< /Length " & Len(contents) & " >>" & vbCrLf & _
            "stream" & vbCrLf & contents & "endstream", Len(text), offsets)
    Next page
    startXref = Len(text)
    text = text & "xref" & vbCrLf & "0 " & (count + 1) & vbCrLf & "0000000000 65535 f" & vbCrLf
    For i = 1 To count
        text = text & Right$("0000000000" & CStr(offsets(i)), 10) & " 00000 n" & vbCrLf
    Next i
    UIQuoteLayoutPDF = text & "trailer" & vbCrLf & "<< /Size " & (count + 1) & " /Root 1 0 R >>" & vbCrLf & _
        "startxref" & vbCrLf & startXref & vbCrLf & "%%EOF" & vbCrLf
End Function

Private Function UIPDFSummary(ByVal result As UIQuoteResult) As String
    Dim text As String
    Dim totalParts As Long
    Dim pairs As Long
    Dim rotated As Long
    Dim i As Long
    Dim stockArea As Double
    Dim unusedArea As Double
    Dim navy As String
    Dim gray As String
    navy = "0.122 0.286 0.49"
    gray = "0.35 0.35 0.35"
    For i = 0 To UBound(result.SheetPlacements)
        totalParts = totalParts + result.SheetPlacements(i, 8)
        If result.SheetPlacements(i, 8) = 2 Then pairs = pairs + 1
        If result.SheetPlacements(i, 6) = 90 Then rotated = rotated + 1
    Next i
    stockArea = result.SheetCount * 7200
    unusedArea = stockArea - result.TotalArea
    text = UIPDFHeader("Sheet Layout - Optimization Summary", result.Material & " | Thickness: " & _
        UIInchesText(result.Thickness) & " in | Stock: 120 x 60 in | " & totalParts & " physical parts")
    text = text & UIPDFStyledText(36, 516, 10, "Design: " & result.WallCount & " walls | " & result.Orientation & " | " & _
        result.ReturnLabel & " | L x W x H: " & UIInchesText(result.OverallLength) & " x " & _
        UIInchesText(result.OverallWidth) & " x " & UIInchesText(result.Height) & " in", False, gray)
    If result.Quantity > 1 Or Len(result.EstimateLabel) > 0 Then
        text = text & UIPDFStyledText(36, 502, 9, result.EstimateLabel & " | Quantity: " & result.Quantity & _
            " | Parts, sheets and areas include all copies of this design.", True, navy)
    End If
    text = text & UIPDFCard(36, "SHEETS REQUIRED", CStr(result.SheetCount), _
        "Area-only minimum: " & result.AreaMinimumSheets & " sheets", navy)
    text = text & UIPDFCard(219, "SHEET UTILIZATION", Format$(result.UtilizationPercent, "0.00") & "%", _
        "Profile area: " & Format$(result.TotalArea / 144, "0.00") & " sq ft", navy)
    text = text & UIPDFCard(402, "UNUSED / POTENTIAL SCRAP", Format$(result.UnusedPercent, "0.00") & "%", _
        "Unused area: " & Format$(unusedArea / 144, "0.00") & " sq ft", "0.749 0.341 0")
    text = text & UIPDFCard(585, "PURCHASED STOCK", Format$(stockArea / 144, "0.00") & " sq ft", _
        result.SheetCount & " sheets of 120 x 60 in", navy)
    text = text & UIPDFRect(36, 386, 720, 12, "0.976 0.843 0.718")
    text = text & UIPDFRect(36, 386, 720 * result.UtilizationPercent / 100, 12, navy)
    text = text & UIPDFStyledText(36, 372, 9, "Used for profiles: " & Format$(result.UtilizationPercent, "0.00") & "%", True, navy)
    text = text & UIPDFRightText(756, 372, 9, "Unused / potential scrap: " & Format$(result.UnusedPercent, "0.00") & "%", True, "0.749 0.341 0")
    text = text & UIPDFStyledText(36, 342, 13, "How the layout was optimized", True, navy)
    text = text & UIPDFRect(36, 335, 340, 1, navy)
    text = text & UIPDFStep(314, "1. Fit gussets together", _
        "Gussets are flipped and paired up so less metal is wasted (" & pairs & " pairs).")
    text = text & UIPDFStep(276, "2. Try many arrangements", _
        "Up to " & result.NestTrialCount & " different layouts are tried and the best one is kept.")
    text = text & UIPDFStep(238, "3. Turn parts when it helps", _
        "Parts can be turned sideways to fill gaps (" & rotated & " parts turned).")
    text = text & UIPDFStep(200, "4. Remove a sheet if possible", _
        "If the emptiest sheet's parts fit on other sheets, that sheet is dropped.")
    text = text & UIPDFStyledText(410, 342, 13, "Selected result", True, navy)
    text = text & UIPDFRect(410, 335, 346, 1, navy)
    text = text & UIPDFFact(314, "Method", result.NestMethod)
    text = text & UIPDFFact(294, "32-trial baseline", result.NestBaselineSheets & " sheets")
    text = text & UIPDFFact(274, "Area-only minimum", result.AreaMinimumSheets & " sheets (lower bound)")
    text = text & UIPDFFact(254, "Physical parts", CStr(totalParts))
    text = text & UIPDFFact(234, "Part gap / edge margin", UIInchesText(result.PartGap) & " in / " & UINestMarginText(result))
    text = text & UIPDFFact(214, "Rotated placements", CStr(rotated))
    text = text & UIPDFStyledText(410, 192, 8.5, "Area-only minimum ignores fit, gaps and margins; not a feasible layout.", False, gray)
    If Len(result.NestWarning) > 0 Then
        text = text & UIPDFStyledText(36, 175, 8.5, "NOTE: " & result.NestWarning, True, "0.749 0.341 0")
    End If
    text = text & UIPDFRect(36, 74, 720, 96, "1 0.969 0.89")
    text = text & UIPDFRect(36, 74, 4, 96, "0.749 0.341 0")
    text = text & UIPDFStyledText(50, 152, 11, "Keep in mind", True, "0.6 0.27 0")
    text = text & UIPDFText(50, 134, 9.5, "- Best layout found by these heuristics - not a guaranteed global minimum or the least possible scrap.")
    text = text & UIPDFText(50, 118, 9.5, "- Unused material includes margins, gaps, offcuts and reusable remnants - not all discarded scrap.")
    text = text & UIPDFText(50, 102, 9.5, "- Not modeled: holes/reliefs, kerf, lead-ins, clamps, common-line cuts and remnant reuse.")
    text = text & UIPDFText(50, 86, 9.5, "- Estimator turns parts 0/90 degrees; CypCut can use finer angles. Operations controls production nesting.")
    If result.PartGap = 0 Then
        text = text & UIPDFStyledText(36, 64, 8.5, "ZERO GAP: touching edges only. Operations must set up co-edge/common-line cuts in CypCut.", True, "0.6 0.27 0")
    End If
    text = text & UIPDFStyledText(36, 50, 10, "ESTIMATING LAYOUT - NOT A PRODUCTION CUTTING PLAN. Operations approval required.", True, navy)
    text = text & UIPDFStyledText(36, 34, 9, "Next pages: sheet usage table, parts list with total quantities, then a reduced-scale layout of each sheet.", False, gray)
    UIPDFSummary = text
End Function

Private Function UIPDFUsage(ByVal result As UIQuoteResult, ByVal page As Long, ByVal pages As Long) As String
    Dim rows As Variant
    Dim i As Long
    Dim last As Long
    Dim top As Double
    Dim y As Double
    Dim bold As Boolean
    Dim color As String
    Dim text As String
    Dim navy As String
    navy = "0.122 0.286 0.49"
    rows = UIQuoteSheetStats(result)
    text = UIPDFHeader("Sheet Usage - Unused / Potential Scrap", _
        "120 x 60-inch stock | Usage table " & (page + 1) & " of " & pages & " | Areas in square feet")
    text = text & UIPDFStyledText(36, 512, 10, "Overall: " & result.SheetCount & " sheets | Utilization " & _
        Format$(rows(result.SheetCount, 4), "0.00") & "% | Unused / potential scrap " & _
        Format$(rows(result.SheetCount, 5), "0.00") & "%", True, navy)
    text = text & UIPDFRect(36, 476, 720, 20, navy)
    text = text & UIPDFStyledText(46, 482, 9, "Sheet", True, "1 1 1") & UIPDFRightText(140, 482, 9, "Parts", True, "1 1 1") & _
        UIPDFRightText(240, 482, 9, "Used (ft^2)", True, "1 1 1") & _
        UIPDFRightText(420, 482, 9, "Unused / potential scrap (ft^2)", True, "1 1 1") & _
        UIPDFRightText(500, 482, 9, "Utilization", True, "1 1 1") & _
        UIPDFStyledText(520, 482, 9, "Used vs. unused", True, "1 1 1") & UIPDFRightText(746, 482, 9, "Unused %", True, "1 1 1")
    last = (page + 1) * 24 - 1
    If last > result.SheetCount Then last = result.SheetCount
    For i = page * 24 To last
        top = 476 - (i - page * 24) * 17
        y = top - 12
        bold = (i = result.SheetCount)
        color = "0 0 0"
        If bold Then
            text = text & UIPDFRect(36, top - 17, 720, 17, "0.851 0.89 0.937") & UIPDFRect(36, top - 1, 720, 1, navy)
            color = navy
        ElseIf (i - page * 24) Mod 2 = 1 Then
            text = text & UIPDFRect(36, top - 17, 720, 17, "0.953 0.957 0.965")
        End If
        text = text & UIPDFStyledText(46, y, 10, CStr(rows(i, 0)), bold, color) & _
            UIPDFRightText(140, y, 10, CStr(rows(i, 1)), bold, color) & _
            UIPDFRightText(240, y, 10, Format$(rows(i, 2), "0.00"), bold, color) & _
            UIPDFRightText(420, y, 10, Format$(rows(i, 3), "0.00"), bold, color) & _
            UIPDFRightText(500, y, 10, Format$(rows(i, 4), "0.00") & "%", bold, color) & _
            UIPDFRect(520, top - 12, 150, 7, "0.976 0.843 0.718") & _
            UIPDFRect(520, top - 12, 150 * rows(i, 4) / 100, 7, navy) & _
            UIPDFRightText(746, y, 10, Format$(rows(i, 5), "0.00") & "%", bold, color)
    Next i
    text = text & UIPDFStyledText(36, 48, 9, "Unused includes reusable remnants, margins, part gaps and offcuts; not all of it is discarded scrap.", False, "0.35 0.35 0.35")
    text = text & UIPDFStyledText(36, 34, 9, "Total utilization is area-weighted across full purchased sheets; displayed values are rounded.", False, "0.35 0.35 0.35")
    UIPDFUsage = text
End Function

Private Function UIPDFHeader(ByVal title As String, ByVal subtitle As String) As String
    UIPDFHeader = UIPDFRect(0, 536, 792, 76, "0.122 0.286 0.49") & _
        UIPDFStyledText(36, 576, 22, title, True, "1 1 1") & _
        UIPDFStyledText(36, 552, 11, subtitle, False, "0.8 0.87 0.94")
End Function

Private Function UIPDFCard(ByVal x As Double, ByVal caption As String, ByVal value As String, _
    ByVal detail As String, ByVal color As String) As String
    UIPDFCard = UIPDFRect(x, 418, 171, 72, "0.929 0.949 0.973") & UIPDFRect(x, 418, 4, 72, color) & _
        UIPDFStyledText(x + 14, 472, 8.5, caption, True, "0.35 0.35 0.35") & _
        UIPDFStyledText(x + 14, 450, 22, value, True, color) & _
        UIPDFStyledText(x + 14, 428, 9, detail, False, "0.35 0.35 0.35")
End Function

Private Function UIPDFStep(ByVal y As Double, ByVal title As String, ByVal detail As String) As String
    UIPDFStep = UIPDFStyledText(36, y, 10.5, title, True, "0 0 0") & _
        UIPDFStyledText(50, y - 14, 9.5, detail, False, "0.35 0.35 0.35")
End Function

Private Function UIPDFFact(ByVal y As Double, ByVal label As String, ByVal value As String) As String
    UIPDFFact = UIPDFStyledText(410, y, 9.5, label, False, "0.35 0.35 0.35") & _
        UIPDFStyledText(545, y, 10, value, True, "0 0 0")
End Function

Private Function UIPDFRect(ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal color As String) As String
    UIPDFRect = "q " & color & " rg " & UIInchesText(Round(x, 2)) & " " & UIInchesText(Round(y, 2)) & " " & _
        UIInchesText(Round(w, 2)) & " " & UIInchesText(Round(h, 2)) & " re f Q" & vbCrLf
End Function

Private Function UIPDFRightText(ByVal edge As Double, ByVal y As Double, ByVal size As Double, _
    ByVal label As String, ByVal bold As Boolean, ByVal color As String) As String
    Dim i As Long
    Dim c As String
    Dim width As Double
    ' Approximate Helvetica advance widths (1/1000 em) for right-aligned numbers.
    For i = 1 To Len(label)
        c = Mid$(label, i, 1)
        If InStr("0123456789$", c) > 0 Then
            width = width + 556
        ElseIf InStr(".,: /", c) > 0 Then
            width = width + 278
        ElseIf c = "%" Then
            width = width + 889
        ElseIf c = "(" Or c = ")" Then
            width = width + 333
        Else
            width = width + 600
        End If
    Next i
    UIPDFRightText = UIPDFStyledText(edge - width * size / 1000, y, size, label, bold, color)
End Function

Private Function UIPDFObject(ByVal id As Long, ByVal body As String, ByVal offset As Long, ByRef offsets() As Long) As String
    offsets(id) = offset
    UIPDFObject = id & " 0 obj" & vbCrLf & body & vbCrLf & "endobj" & vbCrLf
End Function

Private Function UIPDFSheet(ByVal result As UIQuoteResult, ByVal sheet As Long) As String
    Dim text As String
    Dim i As Long
    Dim j As Long
    Dim index As Long
    Dim names As Variant
    Dim label As String
    Dim codes As String
    Dim quantities(0 To 7) As Long
    Dim area As Double
    Dim physicalCount As Long
    Dim x As Double
    Dim y As Double
    Dim w As Double
    Dim h As Double
    Dim b As Double
    Dim a As Double
    Dim shift As Double
    Dim points As Variant
    Dim paired As Boolean
    Dim gusset As Boolean
    text = UIPDFText(36, 576, 18, "Estimated Sheet Layout - Sheet " & sheet & " of " & result.SheetCount)
    text = text & UIPDFText(36, 556, 10, result.Material & " | Thickness: " & UIInchesText(result.Thickness) & _
        " in | Stock: 120 x 60 in | Gap: " & UIInchesText(result.PartGap) & " in | Edge margin: " & UINestMarginText(result))
    text = text & "0.5 w 0 G" & vbCrLf & UIPDFOutline(Array(Array(0, 0), Array(120, 0), Array(120, 60), Array(0, 60)), 0, 0, 0, 60, False)
    text = text & "[3 3] 0 d 0.5 G" & vbCrLf & UIPDFOutline(Array(Array(result.EdgeMarginLength, result.EdgeMarginWidth), _
        Array(120 - result.EdgeMarginLength, result.EdgeMarginWidth), Array(120 - result.EdgeMarginLength, 60 - result.EdgeMarginWidth), _
        Array(result.EdgeMarginLength, 60 - result.EdgeMarginWidth)), 0, 0, 0, 60, False)
    text = text & "[] 0 d 0 G 0.3 w 0.92 g" & vbCrLf
    b = result.GussetBottomWidth
    a = result.GussetTopWidth
    h = result.GussetHeight
    shift = result.PartGap * Sqr(1 + ((b - a) / h) ^ 2)
    For i = 0 To UBound(result.SheetPlacements)
        If result.SheetPlacements(i, 0) = sheet Then
            area = area + result.SheetPlacements(i, 7)
            physicalCount = physicalCount + result.SheetPlacements(i, 8)
            x = result.SheetPlacements(i, 2)
            y = result.SheetPlacements(i, 3)
            paired = (result.SheetPlacements(i, 8) = 2)
            label = Replace(CStr(result.SheetPlacements(i, 1)), " (opposed pair)", "")
            names = Split(label, " + ")
            codes = ""
            For j = 0 To UBound(names)
                index = UIPDFComponentIndex(result, CStr(names(j)))
                quantities(index) = quantities(index) + 1
                If j > 0 Then codes = codes & "+"
                codes = codes & UIPDFPartLabel(index)
            Next j
            gusset = result.Components(index).TopWidth > 0
            If gusset Then
                points = Array(Array(0, 0), Array(b, 0), Array(a, h), Array(0, h))
                text = text & UIPDFOutline(points, x, y, result.SheetPlacements(i, 6), h, True)
                If paired Then
                    points = Array(Array(b + shift, 0), Array(b + a + shift, 0), _
                        Array(b + a + shift, h), Array(a + shift, h))
                    text = text & UIPDFOutline(points, x, y, result.SheetPlacements(i, 6), h, True)
                End If
            Else
                w = result.SheetPlacements(i, 4)
                h = result.SheetPlacements(i, 5)
                text = text & UIPDFOutline(Array(Array(0, 0), Array(w, 0), Array(w, h), Array(0, h)), x, y, 0, h, True)
            End If
            h = result.GussetHeight
            text = text & UIPDFText(36 + x * 6 + 2, 190 + y * 6 + 2, 5, codes)
        End If
    Next i
    text = text & UIPDFText(36, 170, 11, "This sheet: " & physicalCount & " parts | Utilization: " & _
        Format$(area / 7200 * 100, "0.00") & "% | Unused: " & Format$(100 - area / 7200 * 100, "0.00") & "%")
    text = text & UIPDFStyledText(36, 154, 9, "PART LEGEND  (parts on this sheet)", True, "0.35 0.35 0.35")
    j = 0
    For i = 0 To UBound(result.Components)
        If quantities(i) > 0 Then
            x = 36 + (j \ 4) * 360
            y = 140 - (j Mod 4) * 12
            text = text & UIPDFStyledText(x, y, 9, UIPDFPartLabel(i), True, "0.122 0.286 0.49")
            text = text & UIPDFText(x + 32, y, 9, result.Components(i).Name & ": " & quantities(i))
            j = j + 1
        End If
    Next i
    text = text & UIPDFText(36, 74, 10, "Overall: " & result.SheetCount & " sheets | Utilization: " & _
        Format$(result.UtilizationPercent, "0.00") & "% | Unused: " & Format$(result.UnusedPercent, "0.00") & "%")
    text = text & UIPDFText(36, 54, 10, "ESTIMATING LAYOUT - NOT A PRODUCTION CUTTING PLAN. Operations approval required.")
    text = text & UIPDFText(36, 38, 9, "Not full size. Profiles exclude holes/reliefs; no kerf or toolpaths. Unused includes reusable remnants.")
    UIPDFSheet = text
End Function

Private Function UIPDFParts(ByVal result As UIQuoteResult) As String
    Dim counts As Variant
    Dim part As UIQuoteComponent
    Dim i As Long
    Dim row As Long
    Dim top As Double
    Dim y As Double
    Dim totalQty As Long
    Dim totalArea As Double
    Dim size As String
    Dim text As String
    Dim navy As String
    navy = "0.122 0.286 0.49"
    counts = UIPDFPartCounts(result)
    text = UIPDFHeader("Parts List - Total Quantities", result.Material & " | Thickness: " & _
        UIInchesText(result.Thickness) & " in | Parts used across all " & result.SheetCount & " sheets")
    text = text & UIPDFRect(36, 476, 720, 20, navy)
    text = text & UIPDFStyledText(46, 482, 9, "Label", True, "1 1 1") & _
        UIPDFStyledText(100, 482, 9, "Description", True, "1 1 1") & _
        UIPDFStyledText(260, 482, 9, "Shape", True, "1 1 1") & _
        UIPDFStyledText(340, 482, 9, "Profile size (in)", True, "1 1 1") & _
        UIPDFRightText(640, 482, 9, "Total qty", True, "1 1 1") & _
        UIPDFRightText(746, 482, 9, "Total area (ft^2)", True, "1 1 1")
    For i = 0 To UBound(result.Components)
        If counts(i) > 0 Then
            Set part = result.Components(i)
            top = 476 - row * 20
            y = top - 14
            If row Mod 2 = 1 Then text = text & UIPDFRect(36, top - 20, 720, 20, "0.953 0.957 0.965")
            If part.TopWidth > 0 Then
                size = Format$(part.Length, "0.00") & " H x " & Format$(part.Width, "0.00") & " bottom / " & _
                    Format$(part.TopWidth, "0.00") & " top"
            Else
                size = Format$(part.Length, "0.00") & " x " & Format$(part.Width, "0.00")
            End If
            text = text & UIPDFStyledText(46, y, 10, UIPDFPartLabel(i), True, navy) & _
                UIPDFStyledText(100, y, 10, part.Name, False, "0 0 0") & _
                UIPDFStyledText(260, y, 10, part.Shape, False, "0 0 0") & _
                UIPDFStyledText(340, y, 10, size, False, "0 0 0") & _
                UIPDFRightText(640, y, 10, CStr(counts(i)), True, "0 0 0") & _
                UIPDFRightText(746, y, 10, Format$(part.Area * counts(i) / 144, "0.00"), False, "0 0 0")
            totalQty = totalQty + counts(i)
            totalArea = totalArea + part.Area * counts(i)
            row = row + 1
        End If
    Next i
    top = 476 - row * 20
    y = top - 14
    text = text & UIPDFRect(36, top - 20, 720, 20, "0.851 0.89 0.937") & UIPDFRect(36, top - 1, 720, 1, navy)
    text = text & UIPDFStyledText(46, y, 10, "TOTAL", True, navy) & _
        UIPDFRightText(640, y, 10, CStr(totalQty), True, navy) & _
        UIPDFRightText(746, y, 10, Format$(totalArea / 144, "0.00"), True, navy)
    text = text & UIPDFStyledText(36, 48, 9, "Only parts used in this layout are listed. Gusset pairs count as two parts. Labels match the sheet pages.", False, "0.35 0.35 0.35")
    text = text & UIPDFStyledText(36, 34, 9, "Profile areas exclude holes/reliefs; quantities are estimating values, not a production order.", False, "0.35 0.35 0.35")
    UIPDFParts = text
End Function

Private Function UIPDFPartCounts(ByVal result As UIQuoteResult) As Variant
    Dim counts() As Long
    Dim names As Variant
    Dim i As Long
    Dim j As Long
    Dim index As Long
    ReDim counts(0 To UBound(result.Components))
    For i = 0 To UBound(result.SheetPlacements)
        names = Split(Replace(CStr(result.SheetPlacements(i, 1)), " (opposed pair)", ""), " + ")
        For j = 0 To UBound(names)
            index = UIPDFComponentIndex(result, CStr(names(j)))
            counts(index) = counts(index) + 1
        Next j
    Next i
    UIPDFPartCounts = counts
End Function

Private Function UIPDFPartLabel(ByVal index As Long) As String
    Dim labels As Variant
    labels = Array("CW", "CBP", "SW", "SBP", "LW", "RW", "MG", "CG")
    UIPDFPartLabel = labels(index)
End Function

Private Function UIPDFComponentIndex(ByVal result As UIQuoteResult, ByVal name As String) As Long
    Dim i As Long
    For i = 0 To UBound(result.Components)
        If result.Components(i).Name = name Then
            UIPDFComponentIndex = i
            Exit Function
        End If
    Next i
    Err.Raise vbObjectError + 2460, "UIPDFComponentIndex", "Unknown layout component: " & name
End Function

Private Function UIPDFOutline(ByVal points As Variant, ByVal x As Double, ByVal y As Double, _
    ByVal rotation As Long, ByVal height As Double, ByVal filled As Boolean) As String
    Dim i As Long
    Dim xx As Double
    Dim yy As Double
    Dim width As Double
    Dim text As String
    For i = 0 To UBound(points)
        If points(i)(0) > width Then width = points(i)(0)
    Next i
    For i = 0 To UBound(points)
        xx = points(i)(0)
        yy = points(i)(1)
        If rotation = 90 Then
            xx = height - points(i)(1)
            yy = points(i)(0)
        ElseIf rotation = 180 Then
            xx = width - points(i)(0)
            yy = height - points(i)(1)
        ElseIf rotation = 270 Then
            xx = points(i)(1)
            yy = width - points(i)(0)
        End If
        text = text & UIInchesText(36 + (x + xx) * 6) & " " & UIInchesText(190 + (y + yy) * 6)
        If i = 0 Then
            text = text & " m" & vbCrLf
        Else
            text = text & " l" & vbCrLf
        End If
    Next i
    If filled Then
        text = text & "h B" & vbCrLf
    Else
        text = text & "h S" & vbCrLf
    End If
    UIPDFOutline = text
End Function

Private Function UIPDFText(ByVal x As Double, ByVal y As Double, ByVal size As Double, ByVal label As String) As String
    UIPDFText = UIPDFStyledText(x, y, size, label, False, "0 0 0")
End Function

Private Function UIPDFStyledText(ByVal x As Double, ByVal y As Double, ByVal size As Double, ByVal label As String, _
    ByVal bold As Boolean, ByVal color As String) As String
    Dim i As Long
    Dim font As String
    ' Byte offsets must match the ASCII bytes written by the file stream.
    For i = 1 To Len(label)
        If AscW(Mid$(label, i, 1)) < 32 Or AscW(Mid$(label, i, 1)) > 126 Then
            Err.Raise vbObjectError + 2460, "UIPDFText", "PDF labels must contain printable ASCII characters."
        End If
    Next i
    label = Replace(Replace(Replace(label, "\", "\\"), "(", "\("), ")", "\)")
    font = "/F1 "
    If bold Then font = "/F2 "
    UIPDFStyledText = "q " & color & " rg BT " & font & UIInchesText(size) & " Tf " & UIInchesText(Round(x, 2)) & " " & _
        UIInchesText(Round(y, 2)) & " Td (" & label & ") Tj ET Q" & vbCrLf
End Function

Public Sub UISaveQuoteLayoutPDF(ByVal result As UIQuoteResult, ByVal path As String)
    Dim contents As String
    Dim fso As Object
    Dim file As Object
    Dim failure As String
    contents = UIQuoteLayoutPDF(result)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.CreateTextFile(path, False, False)
    On Error GoTo Failed
    file.Write contents
    file.Close
    Exit Sub
Failed:
    failure = Err.Description
    Resume Cleanup
Cleanup:
    On Error GoTo CloseFailed
    file.Close
    On Error GoTo 0
    Err.Raise vbObjectError + 2460, "UISaveQuoteLayoutPDF", _
        "PDF writing failed; the file may be incomplete: " & path & vbCrLf & failure
CloseFailed:
    failure = failure & vbCrLf & "File cleanup also failed: " & Err.Description
    On Error GoTo 0
    Err.Raise vbObjectError + 2460, "UISaveQuoteLayoutPDF", _
        "PDF writing failed; the file may be incomplete: " & path & vbCrLf & failure
End Sub
