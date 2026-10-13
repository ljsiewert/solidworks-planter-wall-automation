Option Explicit

' Code-behind source: paste into a new UserForm named UserForm_QuoteReport.
Private reportBox As MSForms.TextBox
Private headlineCaption As MSForms.Label
Private headlineValue As MSForms.Label
Private designLabel As MSForms.Label
Private tileCaptions(0 To 4) As MSForms.Label
Private tileValues(0 To 4) As MSForms.Label
Private summaryControls As Collection
Private componentsBox As MSForms.Frame
Private gussetsBox As MSForms.Frame
Private totalsBox As MSForms.ListBox
Private sheetBox As MSForms.Frame
Private sheetTitle As MSForms.Label
Private columnLabels As Collection
Private gussetHeadings As Collection
Private sheetHeadings As Collection
Private showingDetails As Boolean
Private quoteResult As UIQuoteResult
Private WithEvents pdfButton As MSForms.CommandButton
Private WithEvents detailsButton As MSForms.CommandButton
Private WithEvents closeButton As MSForms.CommandButton

Private Sub UserForm_Initialize()
    Dim note As MSForms.Label
    Dim heading As MSForms.Label
    Dim captions As Variant
    Dim positions As Variant
    Dim widths As Variant
    Dim i As Long
    Me.Caption = "Planter Material Estimate"
    Me.Width = 780
    Me.Height = 570
    Me.ScrollBars = fmScrollBarsNone
    Set summaryControls = New Collection
    AddSummaryLabel "lblHeaderBand", 12, 10, 744, 66, 8, False, RGB(31, 73, 125), RGB(255, 255, 255), True
    Set headlineCaption = AddSummaryLabel("lblHeadlineCaption", 24, 17, 330, 14, 9, True, 0, RGB(205, 222, 240), False)
    Set headlineValue = AddSummaryLabel("lblHeadlineValue", 24, 32, 330, 38, 22, True, 0, RGB(255, 255, 255), False)
    Set designLabel = AddSummaryLabel("lblDesign", 360, 17, 384, 56, 9, False, 0, RGB(255, 255, 255), False)
    designLabel.TextAlign = fmTextAlignRight
    designLabel.WordWrap = True
    For i = 0 To 4
        AddSummaryLabel "lblTile" & CStr(i), 12 + i * 150.5, 84, 142, 44, 8, False, RGB(237, 242, 248), 0, True
        Set tileCaptions(i) = AddSummaryLabel("lblTileCaption" & CStr(i), 20 + i * 150.5, 89, 128, 12, 8, False, 0, RGB(89, 89, 89), False)
        Set tileValues(i) = AddSummaryLabel("lblTileValue" & CStr(i), 20 + i * 150.5, 103, 128, 22, 14, True, 0, RGB(31, 73, 125), False)
    Next i
    tileValues(4).ForeColor = RGB(191, 87, 0)
    Set columnLabels = New Collection
    Set gussetHeadings = New Collection
    Set sheetHeadings = New Collection
    captions = Array("Component", "Qty", "Length (in)", "Width/height (in)", "Total wt (lb)", "Material cost")
    positions = Array(12, 212, 262, 362, 492, 602)
    widths = Array(200, 50, 100, 130, 110, 150)
    For i = 0 To 5
        Set heading = Me.Controls.Add("Forms.Label.1", "lblColumn" & CStr(i), True)
        heading.Caption = captions(i)
        heading.Left = positions(i)
        heading.Top = 136
        heading.Width = widths(i)
        heading.Height = 16
        heading.Font.Bold = True
        columnLabels.Add heading
    Next i
    Set componentsBox = AddTableFrame("lstComponents", 152)
    captions = Array("Gusset", "Qty", "Height (in)", "Top width (in)", "Base width (in)", "Total wt (lb)", "Material cost")
    positions = Array(12, 192, 242, 332, 422, 512, 622)
    widths = Array(180, 50, 90, 90, 90, 110, 130)
    For i = 0 To 6
        Set heading = Me.Controls.Add("Forms.Label.1", "lblGussetColumn" & CStr(i), True)
        heading.Caption = captions(i)
        heading.Left = positions(i)
        heading.Top = 254
        heading.Width = widths(i)
        heading.Height = 16
        heading.Font.Bold = True
        columnLabels.Add heading
        gussetHeadings.Add heading
    Next i
    Set gussetsBox = AddTableFrame("lstGussets", 270)
    Set totalsBox = Me.Controls.Add("Forms.ListBox.1", "lstMaterialTotal", True)
    With totalsBox
        .Left = 12
        .Top = 318
        .Width = 744
        .Height = 20
        .ColumnCount = 3
        .ColumnWidths = "490;110;120"
        .Font.Bold = True
        .Font.Size = 10
    End With
    Set sheetTitle = Me.Controls.Add("Forms.Label.1", "lblSheetUsage", True)
    With sheetTitle
        .Caption = "SHEET USAGE - UNUSED / POTENTIAL SCRAP (includes reusable remnants)"
        .Left = 12
        .Top = 344
        .Width = 744
        .Height = 16
        .Font.Bold = True
        .ForeColor = RGB(31, 73, 125)
    End With
    captions = Array("Sheet", "Parts", "Used (ft^2)", "Unused / potential scrap (ft^2)", "Utilization", "Unused %")
    positions = Array(12, 92, 152, 272, 492, 602)
    widths = Array(80, 60, 120, 220, 110, 130)
    For i = 0 To 5
        Set heading = Me.Controls.Add("Forms.Label.1", "lblSheetColumn" & CStr(i), True)
        heading.Caption = captions(i)
        heading.Left = positions(i)
        heading.Top = 360
        heading.Width = widths(i)
        heading.Height = 16
        heading.Font.Bold = True
        columnLabels.Add heading
        sheetHeadings.Add heading
    Next i
    Set sheetBox = AddTableFrame("lstSheetUsage", 376)
    Set reportBox = Me.Controls.Add("Forms.TextBox.1", "txtQuoteReport", True)
    With reportBox
        .Left = 12
        .Top = 10
        .Width = 744
        .Height = 494
        .MultiLine = True
        .ScrollBars = fmScrollBarsVertical
        .Locked = True
        .WordWrap = True
        .Visible = False
    End With
    Set note = Me.Controls.Add("Forms.Label.1", "lblEstimateNotes", True)
    note.Left = 12
    note.Top = 462
    note.Width = 744
    note.Height = 46
    note.Font.Size = 8
    note.ForeColor = RGB(89, 89, 89)
    note.Caption = "Material only - not a final manufacturing quote. Sheet offcuts, fabrication, labor and margin excluded." & vbCrLf & _
        "Holes/reliefs not deducted. Flat heights use calibrated measurements; other heights are predictions." & vbCrLf & _
        "Both steels use an estimating density of 7850 kg/m^3; BORCON-specific density is unverified." & vbCrLf & _
        "Sheet estimates use provisional gap/margin settings; Operations must verify. Unused area may be reusable remnant."
    Set detailsButton = Me.Controls.Add("Forms.CommandButton.1", "cmdDetails", True)
    With detailsButton
        .Caption = "Technical Details"
        .Left = 12
        .Top = 512
        .Width = 130
        .Height = 24
    End With
    Set pdfButton = Me.Controls.Add("Forms.CommandButton.1", "cmdLayoutPDF", True)
    With pdfButton
        .Caption = "Preview Sheet Layout"
        .Left = 152
        .Top = 512
        .Width = 150
        .Height = 24
        .Enabled = False
    End With
    Set closeButton = Me.Controls.Add("Forms.CommandButton.1", "cmdClose", True)
    With closeButton
        .Caption = "Close"
        .Left = 666
        .Top = 512
        .Width = 90
        .Height = 24
        .Default = True
        .Cancel = True
    End With
End Sub

Public Sub SetResult(ByVal result As UIQuoteResult)
    Dim cards As Variant
    Dim componentRows As Variant
    Dim gussetRows As Variant
    Dim sheetRows As Variant
    Dim heading As MSForms.Label
    Dim y As Single
    Dim i As Long
    Set quoteResult = result
    pdfButton.Enabled = result.NestAvailable
    cards = UIQuoteSummaryCards(result)
    headlineCaption.Caption = cards(0, 0)
    headlineValue.Caption = cards(0, 1)
    designLabel.Caption = cards(1, 1)
    designLabel.ControlTipText = Replace(cards(1, 1), vbCrLf, "  ")
    If Not result.HasValidTotals Then
        headlineCaption.ForeColor = RGB(255, 214, 153)
        designLabel.TextAlign = fmTextAlignLeft
    End If
    For i = 0 To 4
        tileCaptions(i).Caption = cards(i + 2, 0)
        tileValues(i).Caption = cards(i + 2, 1)
    Next i
    componentRows = UIQuoteComponentRows(result)
    gussetRows = UIQuoteGussetRows(result)
    totalsBox.List = UIQuoteTotalRows(result)
    sheetRows = UIQuoteSheetRows(result)
    sheetTitle.Caption = "SHEET USAGE - UNUSED / POTENTIAL SCRAP (includes reusable remnants)"
    If Not result.NestAvailable Then sheetTitle.Caption = result.NestStatus
    reportBox.Text = UIQuoteReport(result)
    ' Each table is exactly as tall as its rows; controls below move up to follow.
    y = FillTable(componentsBox, componentRows, Array(200, 50, 100, 130, 110, 150), componentsBox.Top, 200) + 6
    For Each heading In gussetHeadings
        heading.Top = y
    Next heading
    y = FillTable(gussetsBox, gussetRows, Array(180, 50, 90, 90, 90, 110, 130), y + 16, 120) + 4
    totalsBox.Top = y
    sheetTitle.Top = y + totalsBox.Height + 6
    For Each heading In sheetHeadings
        heading.Top = sheetTitle.Top + 16
    Next heading
    FillTable sheetBox, sheetRows, Array(80, 60, 120, 220, 110, 130), sheetTitle.Top + 32, 456 - (sheetTitle.Top + 32)
End Sub

Private Function AddTableFrame(ByVal controlName As String, ByVal top As Single) As MSForms.Frame
    Dim table As MSForms.Frame
    Set table = Me.Controls.Add("Forms.Frame.1", controlName, True)
    table.Caption = ""
    table.Left = 12
    table.Top = top
    table.Width = 744
    table.BackColor = RGB(255, 255, 255)
    table.BorderStyle = fmBorderStyleSingle
    table.BorderColor = RGB(171, 171, 171)
    table.SpecialEffect = fmSpecialEffectFlat
    Set AddTableFrame = table
End Function

' Fills a frame with one label per cell. Rows taller than maxHeight scroll inside the frame.
Private Function FillTable(ByVal table As MSForms.Frame, ByVal rows As Variant, ByVal widths As Variant, _
    ByVal top As Single, ByVal maxHeight As Single) As Single
    Const ROW_HEIGHT As Single = 15
    Dim cell As MSForms.Label
    Dim r As Long
    Dim c As Long
    Dim x As Single
    Dim contentHeight As Single
    Dim isTotal As Boolean
    table.Controls.Clear
    For r = 0 To UBound(rows, 1)
        x = 0
        isTotal = (CStr(rows(r, 0)) = "TOTAL")
        For c = 0 To UBound(rows, 2)
            Set cell = table.Controls.Add("Forms.Label.1")
            cell.Left = x
            cell.Top = r * ROW_HEIGHT
            cell.Width = widths(c)
            cell.Height = ROW_HEIGHT
            cell.Caption = " " & CStr(rows(r, c))
            cell.Font.Size = 10
            cell.Font.Bold = isTotal
            cell.WordWrap = False
            cell.BackStyle = fmBackStyleOpaque
            If isTotal Then
                cell.BackColor = RGB(217, 227, 239)
            ElseIf r Mod 2 = 1 Then
                cell.BackColor = RGB(243, 244, 246)
            Else
                cell.BackColor = RGB(255, 255, 255)
            End If
            x = x + widths(c)
        Next c
    Next r
    contentHeight = (UBound(rows, 1) + 1) * ROW_HEIGHT
    table.Top = top
    table.ScrollTop = 0
    If contentHeight + 4 > maxHeight Then
        table.Height = maxHeight
        table.ScrollBars = fmScrollBarsVertical
        table.ScrollHeight = contentHeight
    Else
        table.Height = contentHeight + 4
        table.ScrollBars = fmScrollBarsNone
    End If
    FillTable = table.Top + table.Height
End Function

Private Sub pdfButton_Click()
    On Error GoTo Failed
    UIPreviewQuoteLayoutPDF quoteResult
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Sheet Layout Preview Failed"
End Sub

Private Sub detailsButton_Click()
    Dim heading As MSForms.Label
    Dim item As MSForms.Label
    showingDetails = Not showingDetails
    For Each item In summaryControls
        item.Visible = Not showingDetails
    Next item
    componentsBox.Visible = Not showingDetails
    gussetsBox.Visible = Not showingDetails
    totalsBox.Visible = Not showingDetails
    sheetBox.Visible = Not showingDetails
    sheetTitle.Visible = Not showingDetails
    For Each heading In columnLabels
        heading.Visible = Not showingDetails
    Next heading
    reportBox.Visible = showingDetails
    If showingDetails Then
        detailsButton.Caption = "Back to Summary"
    Else
        detailsButton.Caption = "Technical Details"
    End If
End Sub

Private Sub closeButton_Click()
    Me.Hide
End Sub

Private Function AddSummaryLabel(ByVal controlName As String, ByVal x As Single, ByVal y As Single, _
    ByVal w As Single, ByVal h As Single, ByVal size As Single, ByVal bold As Boolean, _
    ByVal fillColor As Long, ByVal textColor As Long, ByVal opaque As Boolean) As MSForms.Label
    Dim item As MSForms.Label
    Set item = Me.Controls.Add("Forms.Label.1", controlName, True)
    item.Left = x
    item.Top = y
    item.Width = w
    item.Height = h
    item.Font.Size = size
    item.Font.Bold = bold
    item.ForeColor = textColor
    If opaque Then
        item.BackStyle = fmBackStyleOpaque
        item.BackColor = fillColor
    Else
        item.BackStyle = fmBackStyleTransparent
    End If
    summaryControls.Add item
    Set AddSummaryLabel = item
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub
