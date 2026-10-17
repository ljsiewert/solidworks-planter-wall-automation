Option Explicit

' Code-behind source: paste into a new UserForm named UserForm_FieldQuote.
Private WithEvents detailsButton As MSForms.CommandButton
Private WithEvents previewButton As MSForms.CommandButton
Private WithEvents marginBox As MSForms.TextBox
Private WithEvents marginSpin As MSForms.SpinButton
Private WithEvents closeButton As MSForms.CommandButton
Private designs As Collection
Private designList As MSForms.ListBox
Private summaryLabel As MSForms.Label
Private sellCaption As MSForms.Label
Private metricValues(0 To 6) As MSForms.Label
Private previewResult As UIQuoteResult
Private updatingMargin As Boolean

Public Property Get PreviewDesign() As UIQuoteResult
    Set PreviewDesign = previewResult
End Property

Private Sub UserForm_Initialize()
    Dim label As MSForms.Label
    Dim captions As Variant
    Dim widths As Variant
    Dim columnWidths As String
    Dim x As Single
    Dim i As Long
    Me.Caption = "Field Assembly - Planter Estimate"
    Me.Width = 790
    Me.Height = 562
    Me.ScrollBars = fmScrollBarsNone
    Set label = Me.Controls.Add("Forms.Label.1", "lblFieldTitle", True)
    label.Caption = "FIELD ASSEMBLY ESTIMATE"
    label.Left = 12
    label.Top = 12
    label.Width = 750
    label.Height = 24
    label.Font.Size = 14
    label.Font.Bold = True
    label.ForeColor = RGB(31, 73, 125)
    Set metricValues(0) = AddPriceCard("Sell", "FIELD SELL PRICE", 12, 42, 750, 76, 28, True)
    Set sellCaption = Me.Controls("lblSellCaption")
    Set metricValues(1) = AddPriceCard("Profit", "PROFIT", 12, 126, 371, 60, 20, False)
    Set metricValues(2) = AddPriceCard("Cost", "TOTAL COST", 391, 126, 371, 60, 20, False)
    Set metricValues(3) = AddPriceCard("Material", "MATERIAL", 12, 194, 182, 48, 14, False)
    Set metricValues(4) = AddPriceCard("Labor", "LABOR", 202, 194, 182, 48, 14, False)
    Set metricValues(5) = AddPriceCard("Paint", "PAINT", 392, 194, 182, 48, 14, False)
    Set metricValues(6) = AddPriceCard("Sheets", "120 x 60 SHEETS", 582, 194, 180, 48, 14, False)
    captions = Array("Design / material / thickness", "Qty", "Sheets", "Total cost", "Sell price")
    widths = Array(360, 55, 65, 120, 130)
    x = 12
    For i = 0 To UBound(widths)
        Set label = Me.Controls.Add("Forms.Label.1", "lblFieldColumn" & CStr(i), True)
        label.Caption = captions(i)
        label.Left = x + 3
        label.Top = 258
        label.Width = widths(i) - 3
        label.Height = 18
        label.Font.Bold = True
        x = x + widths(i)
        If i > 0 Then columnWidths = columnWidths & ";"
        columnWidths = columnWidths & CStr(widths(i))
    Next i
    Set designList = Me.Controls.Add("Forms.ListBox.1", "lstDesigns", True)
    With designList
        .Left = 12
        .Top = 278
        .Width = 750
        .Height = 120
        .ColumnCount = 5
        .ColumnWidths = columnWidths
    End With
    Set summaryLabel = Me.Controls.Add("Forms.Label.1", "lblFieldSummary", True)
    With summaryLabel
        .Left = 12
        .Top = 406
        .Width = 750
        .Height = 18
        .Font.Size = 10
        .Font.Bold = True
        .ForeColor = RGB(31, 73, 125)
    End With
    Set label = Me.Controls.Add("Forms.Label.1", "lblFieldNotes", True)
    label.Caption = "Each design nests separately; copies of that design share sheets. No sheets are shared between designs." & vbCrLf & _
        "Preview selected opens one unsaved CAD design, not its duplicates. Close without saving before another preview."
    label.Left = 12
    label.Top = 428
    label.Width = 750
    label.Height = 30
    label.Font.Size = 9
    Set detailsButton = Me.Controls.Add("Forms.CommandButton.1", "cmdViewDesign", True)
    detailsButton.Caption = "View selected"
    detailsButton.Left = 12
    detailsButton.Top = 494
    detailsButton.Width = 112
    detailsButton.Height = 24
    Set marginBox = Me.Controls.Add("Forms.TextBox.1", "txtFieldMargin", True)
    Set label = Me.Controls.Add("Forms.Label.1", "lblFieldMargin", True)
    label.Caption = "Gross margin % (all designs)"
    label.Left = 12
    label.Top = 466
    label.Width = 164
    label.Height = 18
    Set previewButton = Me.Controls.Add("Forms.CommandButton.1", "cmdPreviewSelected", True)
    previewButton.Caption = "Preview selected"
    previewButton.Left = 136
    previewButton.Top = 494
    previewButton.Width = 130
    previewButton.Height = 24
    marginBox.Left = 180
    marginBox.Top = 462
    marginBox.Width = 50
    marginBox.Height = 22
    marginBox.ControlTipText = "Gross margin percentage (0-95) to apply to every design."
    Set marginSpin = Me.Controls.Add("Forms.SpinButton.1", "spnFieldMargin", True)
    marginSpin.Left = 232
    marginSpin.Top = 462
    marginSpin.Width = 18
    marginSpin.Height = 22
    marginSpin.Min = 0
    marginSpin.Max = 95
    marginSpin.SmallChange = 1
    Set closeButton = Me.Controls.Add("Forms.CommandButton.1", "cmdClose", True)
    closeButton.Caption = "Close"
    closeButton.Left = 672
    closeButton.Top = 494
    closeButton.Width = 90
    closeButton.Height = 24
    closeButton.Cancel = True
End Sub

Private Function AddPriceCard(ByVal name As String, ByVal caption As String, ByVal left As Single, _
    ByVal top As Single, ByVal width As Single, ByVal height As Single, ByVal size As Single, _
    ByVal primary As Boolean) As MSForms.Label
    Dim background As MSForms.Label
    Dim heading As MSForms.Label
    Dim value As MSForms.Label
    Set background = Me.Controls.Add("Forms.Label.1", "lbl" & name & "Card", True)
    background.Left = left
    background.Top = top
    background.Width = width
    background.Height = height
    background.BackStyle = fmBackStyleOpaque
    background.BackColor = IIf(primary, RGB(31, 73, 125), RGB(230, 237, 245))
    Set heading = Me.Controls.Add("Forms.Label.1", "lbl" & name & "Caption", True)
    heading.Caption = caption
    heading.Left = left + 12
    heading.Top = top + 6
    heading.Width = width - 24
    heading.Height = 16
    heading.BackStyle = fmBackStyleTransparent
    heading.Font.Size = 9
    heading.Font.Bold = True
    heading.ForeColor = IIf(primary, RGB(224, 235, 248), RGB(89, 89, 89))
    Set value = Me.Controls.Add("Forms.Label.1", "lbl" & name & "Value", True)
    value.Left = left + 12
    value.Top = top + 23
    value.Width = width - 24
    value.Height = height - 25
    value.BackStyle = fmBackStyleTransparent
    value.Font.Size = size
    value.Font.Bold = True
    value.ForeColor = IIf(primary, RGB(255, 255, 255), RGB(31, 73, 125))
    Set AddPriceCard = value
End Function

Public Sub SetDesigns(ByVal values As Collection)
    Dim first As UIQuoteResult
    If values.Count = 0 Then Err.Raise vbObjectError + 2400, , "Add at least one planter design."
    Set designs = values
    Set first = designs(1)
    UIFieldSetMargin designs, first.MarginPercent
    updatingMargin = True
    marginBox.Text = CStr(first.MarginPercent)
    marginSpin.Value = Int(first.MarginPercent)
    updatingMargin = False
    RefreshTotals
End Sub

Private Sub RefreshTotals()
    Dim result As UIQuoteResult
    Dim totals As Variant
    Dim row As Long
    Dim selected As Long
    Dim indices As Variant
    Dim i As Long
    selected = designList.ListIndex
    designList.Clear
    For Each result In designs
        designList.AddItem result.EstimateLabel & " | " & result.Material & " | " & UIInchesText(result.Thickness) & " in"
        row = designList.ListCount - 1
        designList.List(row, 1) = CStr(result.Quantity)
        If result.HasPricing Then
            designList.List(row, 2) = CStr(result.SheetCount)
            designList.List(row, 3) = "$" & FormatNumber(result.TotalCost, 2)
            designList.List(row, 4) = "$" & FormatNumber(result.SellPrice, 2)
        Else
            designList.List(row, 2) = "Review"
            designList.List(row, 3) = "Withheld"
            designList.List(row, 4) = "Withheld"
        End If
    Next result
    If selected < 0 Then selected = 0
    designList.ListIndex = selected
    totals = UIFieldQuoteTotals(designs)
    summaryLabel.Caption = designs.Count & " designs | " & totals(0) & " planter copies"
    sellCaption.Caption = "FIELD SELL PRICE"
    indices = Array(7, 6, 5, 2, 3, 4, 1)
    If totals(8) Then
        For i = 0 To 5
            metricValues(i).Caption = "$" & FormatNumber(totals(indices(i)), 2)
        Next i
        metricValues(6).Caption = CStr(totals(1))
    Else
        For i = 0 To 6
            metricValues(i).Caption = "Withheld"
        Next i
        sellCaption.Caption = "FIELD TOTALS WITHHELD - VIEW EACH REVIEW ROW FOR THE REASON"
    End If
End Sub

Private Sub detailsButton_Click()
    Dim result As UIQuoteResult
    On Error GoTo Failed
    If Not ApplyMargin() Then Exit Sub
    If designList.ListIndex < 0 Then
        MsgBox "Select a planter design first.", vbExclamation, "Field Assembly"
        Exit Sub
    End If
    Set result = designs(designList.ListIndex + 1)
    UIShowQuote result, False
    RefreshTotals
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Field Assembly"
End Sub

Private Function ApplyMargin() As Boolean
    Dim margin As Double
    If designs Is Nothing Then Exit Function
    On Error GoTo InvalidMargin
    margin = UIQuoteParseRate(marginBox.Text, "Gross margin")
    UIFieldSetMargin designs, margin
    On Error GoTo 0
    marginBox.ForeColor = RGB(0, 0, 0)
    updatingMargin = True
    marginSpin.Value = Int(margin)
    updatingMargin = False
    RefreshTotals
    ApplyMargin = True
    Exit Function
InvalidMargin:
    marginBox.ForeColor = RGB(192, 0, 0)
    sellCaption.Caption = "ENTER A MARGIN FROM 0 TO 95% - LAST VALID PRICE SHOWN"
End Function

Private Sub marginBox_Change()
    If updatingMargin Then Exit Sub
    ApplyMargin
End Sub

Private Sub marginSpin_Change()
    If updatingMargin Then Exit Sub
    updatingMargin = True
    marginBox.Text = CStr(marginSpin.Value)
    updatingMargin = False
    ApplyMargin
End Sub

Private Sub previewButton_Click()
    If Not ApplyMargin() Then Exit Sub
    If designList.ListIndex < 0 Then
        MsgBox "Select a planter design first.", vbExclamation, "Field Assembly"
        Exit Sub
    End If
    Set previewResult = designs(designList.ListIndex + 1)
    Me.Hide
End Sub

Private Sub UserForm_Activate()
    Set previewResult = Nothing
End Sub

Private Sub closeButton_Click()
    Set previewResult = Nothing
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        closeButton_Click
    End If
End Sub
